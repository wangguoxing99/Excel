import os
import json
import random
import logging
import uuid
import re
import io
import datetime
import pandas as pd
import click
from functools import wraps
from flask import Flask, request, render_template, send_file, jsonify, session, redirect, url_for, send_from_directory
from werkzeug.security import generate_password_hash, check_password_hash
from urllib.parse import quote
from datetime import timedelta

app = Flask(__name__)

# --- 配置区域 ---
app.secret_key = os.environ.get("FLASK_SECRET_KEY", os.urandom(24))
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(minutes=30)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['RESULT_FOLDER'] = 'results'
AUTH_FILE = 'auth.json'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)

# --- 日志配置 (用于比对系统) ---
log_stream = []
class WebLogHandler(logging.Handler):
    def emit(self, record):
        log_entry = self.format(record)
        log_stream.append(log_entry)
        if len(log_stream) > 100: log_stream.pop(0)

logger = logging.getLogger('web_logger')
logger.setLevel(logging.INFO)
handler = WebLogHandler()
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M:%S'))
logger.addHandler(handler)

# --- 认证核心逻辑 (升级版) ---

def load_auth_db():
    """读取认证数据库，支持自动迁移旧格式"""
    default_db = {"users": {}}
    if not os.path.exists(AUTH_FILE):
        return default_db
    try:
        with open(AUTH_FILE, 'r') as f:
            data = json.load(f)
            # 兼容迁移：如果是旧的单用户格式，转换为多用户格式
            if "username" in data and "password_hash" in data:
                new_db = {"users": {data["username"]: data["password_hash"]}}
                save_auth_db(new_db) # 立即保存新格式
                return new_db
            return data
    except:
        return default_db

def save_auth_db(data):
    """保存认证数据库"""
    with open(AUTH_FILE, 'w') as f:
        json.dump(data, f)

def add_user_logic(username, password):
    """添加或更新用户"""
    db = load_auth_db()
    db["users"][username] = generate_password_hash(password)
    save_auth_db(db)

def del_user_logic(username):
    """删除用户"""
    db = load_auth_db()
    if username in db["users"]:
        del db["users"][username]
        save_auth_db(db)
        return True
    return False

# --- CLI 命令 (增强版) ---

@app.cli.command("add-user")
@click.argument("username")
@click.argument("password")
def add_user_command(username, password):
    """添加或更新用户: flask add-user <user> <pwd>"""
    add_user_logic(username, password)
    click.echo(f"✅ 用户已更新/添加: {username}")

@app.cli.command("del-user")
@click.argument("username")
def del_user_command(username):
    """删除用户: flask del-user <user>"""
    if del_user_logic(username):
        click.echo(f"🗑️ 用户已删除: {username}")
    else:
        click.echo(f"⚠️ 用户不存在: {username}")

@app.cli.command("list-users")
def list_users_command():
    """列出所有用户"""
    db = load_auth_db()
    users = list(db["users"].keys())
    click.echo(f"👥 当前用户列表: {', '.join(users)}")

# --- 中间件 & 权限控制 ---

@app.before_request
def auth_middleware():
    if request.path.startswith('/static'): return
    
    # 检查是否已初始化（至少有一个用户）
    db = load_auth_db()
    if not db["users"]:
        if request.endpoint != 'setup': return redirect(url_for('setup'))
        return
    
    # 如果已初始化但访问 setup，跳转登录
    if request.endpoint == 'setup': return redirect(url_for('login'))
    
    if request.endpoint in ['login', 'logout']: return
    
    session.permanent = True
    if not session.get('logged_in'): return redirect(url_for('login'))

# --- 基础路由 ---

@app.route('/setup', methods=['GET', 'POST'])
def setup():
    """初始化第一个管理员账号"""
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username and password:
            add_user_logic(username, password)
            return redirect(url_for('login'))
    return render_template('setup.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        db = load_auth_db()
        user_hash = db["users"].get(username)
        
        if user_hash and check_password_hash(user_hash, password):
            session['logged_in'] = True
            session['user'] = username
            return redirect(url_for('portal'))
        else:
            error = "用户名或密码错误"
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
def portal():
    return render_template('portal.html')

# ==========================================
# 模块 A: Excel 智能拆分系统 (Splitter)
# ==========================================

@app.route('/tool/splitter')
def splitter_ui():
    return render_template('splitter.html')

def split_smart_algo(total_qty, days, is_int):
    if days <= 1: return [total_qty]
    amounts = []
    if is_int:
        total_int = int(total_qty)
        if total_int < days: return [1] * total_int
        base = total_int // days
        remainder = total_int % days
        amounts = [base] * days
        indices = list(range(days))
        random.shuffle(indices)
        for i in range(remainder): amounts[indices[i]] += 1
    else:
        weights = [random.uniform(0.8, 1.2) for _ in range(days)]
        sum_weights = sum(weights)
        current_sum = 0
        for w in weights[:-1]:
            val = round((w / sum_weights) * total_qty, 1)
            if val == 0 and total_qty > 1: val = 0.1
            amounts.append(val)
            current_sum += val
        amounts.append(round(total_qty - current_sum, 1))
    return amounts

@app.route('/api/splitter/analyze', methods=['POST'])
def splitter_analyze():
    file = request.files.get('file')
    if not file: return jsonify({"error": "未找到文件"}), 400
    try:
        file_content = file.read()
        file_bytes = io.BytesIO(file_content)
        visible_sheets = []
        filename = file.filename.lower() if file.filename else ""
        if filename.endswith('.xlsx'):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(file_bytes, read_only=True)
                for sheet in wb.worksheets:
                    if sheet.sheet_state == 'visible': visible_sheets.append(sheet.title)
                wb.close()
            except: visible_sheets = []
        file_bytes.seek(0)
        xl = pd.ExcelFile(file_bytes)
        all_sheets = xl.sheet_names
        final_sheets = [s for s in all_sheets if s in visible_sheets] if visible_sheets else all_sheets
        if not final_sheets: final_sheets = all_sheets
        df = xl.parse(final_sheets[0], nrows=10)
        return jsonify({"sheets": final_sheets, "columns": df.columns.tolist()})
    except Exception as e:
        return jsonify({"error": f"解析错误: {str(e)}"}), 500

@app.route('/api/splitter/sheet_info', methods=['POST'])
def splitter_sheet_info():
    file = request.files.get('file')
    sheet_name = request.form.get('sheet_name')
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        columns = df.columns.tolist()
        units = []
        unit_col = next((c for c in columns if '单位' in str(c)), None)
        if unit_col: units = df[unit_col].dropna().unique().tolist()
        return jsonify({"columns": columns, "units": [str(u) for u in units]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/splitter/process', methods=['POST'])
def splitter_process():
    file = request.files.get('file')
    sheet_name = request.form.get('sheet_name')
    target_qty_col = request.form.get('target_qty_col')
    total_days = int(request.form.get('days', 12))
    selected_cols = request.form.getlist('cols')[:10]
    int_units = request.form.getlist('int_units')
    
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        if target_qty_col not in df.columns: return jsonify({"success": False, "error": "指定的数量列不存在"}), 400

        name_col = next((c for c in df.columns if '名称' in str(c)), selected_cols[0] if selected_cols else df.columns[0])
        unit_col = next((c for c in df.columns if '单位' in str(c)), None)
        price_col = next((c for c in df.columns if '单价' in str(c)), None)
        
        df = df.dropna(subset=[target_qty_col])
        df[target_qty_col] = pd.to_numeric(df[target_qty_col], errors='coerce')
        df = df[df[target_qty_col] > 0]
        
        daily_rows = [[] for _ in range(total_days)]
        
        for _, row in df.iterrows():
            unit = str(row.get(unit_col, '')) if unit_col else ""
            qty = row[target_qty_col]
            if pd.isna(qty): continue
            
            if qty <= 3: active_days = 1
            elif qty <= 10: active_days = random.randint(2, min(4, total_days))
            else: active_days = random.randint(3, min(total_days, 10))
            
            is_int = unit in int_units
            splits = split_smart_algo(qty, active_days, is_int)
            days_indices = sorted(random.sample(range(total_days), len(splits)))
            
            for i, day_idx in enumerate(days_indices):
                new_row = {col: row[col] for col in selected_cols if col in row}
                if target_qty_col in new_row: new_row[target_qty_col] = splits[i]
                if price_col in row and '含税金额' in selected_cols:
                    try: new_row['含税金额'] = round(float(row[price_col]) * splits[i], 2)
                    except: pass
                daily_rows[day_idx].append(new_row)
        
        filename = f"拆分_{sheet_name}_{uuid.uuid4().hex[:8]}.xlsx"
        filepath = os.path.join(app.config['RESULT_FOLDER'], filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for i in range(total_days):
                pd.DataFrame(daily_rows[i]).to_excel(writer, sheet_name=f'第{i+1}天', index=False)
        
        return jsonify({"success": True, "filename": filename})
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/splitter/download/<filename>')
def splitter_download(filename):
    return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)

# ==========================================
# 模块 B: 进销项比对系统 (Comparator)
# ==========================================

@app.route('/tool/compare')
def compare_ui():
    return render_template('compare.html')

def clean_name_algo(text):
    if pd.isna(text): return ""
    return re.sub(r'\*.*?\*', '', str(text)).strip()

@app.route('/api/compare/get_logs')
def compare_get_logs():
    global log_stream
    logs = list(log_stream)
    log_stream.clear()
    return jsonify(logs)

@app.route('/api/compare/get_headers', methods=['POST'])
def compare_get_headers():
    f = request.files.get('file')
    if not f: return jsonify({})
    try:
        df = pd.read_excel(f, nrows=1)
        return jsonify({"columns": df.columns.tolist()})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/api/compare/process', methods=['POST'])
def compare_process():
    try:
        f_in = request.files['file_in']
        f_out = request.files['file_out']
        m = request.form
        
        df_in = pd.read_excel(f_in)
        df_out = pd.read_excel(f_out)

        df_in['__key__'] = df_in[m['map_in_name']].apply(clean_name_algo)
        df_out['__key__'] = df_out[m['map_out_name']].apply(clean_name_algo)
        
        in_agg = df_in.groupby('__key__')[[m['map_in_qty'], m['map_in_val']]].sum().reset_index()
        out_agg = df_out.groupby('__key__')[[m['map_out_qty'], m['map_out_val']]].sum().reset_index()

        in_agg.columns = ['关联名称', '进项_数量', '进项_金额']
        out_agg.columns = ['关联名称', '销项_数量', '销项_金额']

        merged = pd.merge(in_agg, out_agg, on='关联名称', how='outer').fillna(0)
        merged['数量差异(销-进)'] = merged['销项_数量'] - merged['进项_数量']
        merged['金额差异(销-进)'] = merged['销项_金额'] - merged['进项_金额']

        res_name = f"result_{uuid.uuid4().hex}.xlsx"
        merged.to_excel(os.path.join(app.config['RESULT_FOLDER'], res_name), index=False)
        
        return jsonify({"success": True, "filename": res_name})
    except Exception as e:
        logger.error(f"处理失败: {str(e)}")
        return jsonify({"success": False, "message": str(e)})

@app.route('/api/compare/download/<filename>')
def compare_download(filename):
    display_name = "进销项比对报告.xlsx"
    response = send_from_directory(app.config['RESULT_FOLDER'], filename)
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(display_name)}"
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
