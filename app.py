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

# --- 日志配置 ---
log_stream = []
class WebLogHandler(logging.Handler):
    def emit(self, record):
        log_stream.append(self.format(record))
        if len(log_stream) > 100: log_stream.pop(0)

logger = logging.getLogger('web_logger')
logger.setLevel(logging.INFO)
handler = WebLogHandler()
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M:%S'))
logger.addHandler(handler)

# --- 认证逻辑 ---
def load_auth_db():
    default_db = {"users": {}}
    if not os.path.exists(AUTH_FILE): return default_db
    try:
        with open(AUTH_FILE, 'r') as f:
            data = json.load(f)
            if "username" in data: # 兼容旧格式
                new_db = {"users": {data["username"]: data["password_hash"]}}
                save_auth_db(new_db)
                return new_db
            return data
    except: return default_db

def save_auth_db(data):
    with open(AUTH_FILE, 'w') as f: json.dump(data, f)

def add_user_logic(username, password):
    db = load_auth_db()
    db["users"][username] = generate_password_hash(password)
    save_auth_db(db)

def del_user_logic(username):
    db = load_auth_db()
    if username in db["users"]:
        del db["users"][username]
        save_auth_db(db)
        return True
    return False

# --- CLI Commands ---
@app.cli.command("add-user")
@click.argument("username")
@click.argument("password")
def add_user_command(username, password):
    add_user_logic(username, password)
    click.echo(f"✅ 用户已更新/添加: {username}")

@app.cli.command("del-user")
@click.argument("username")
def del_user_command(username):
    if del_user_logic(username): click.echo(f"🗑️ 用户已删除: {username}")
    else: click.echo(f"⚠️ 用户不存在: {username}")

@app.cli.command("list-users")
def list_users_command():
    db = load_auth_db()
    click.echo(f"👥 用户列表: {', '.join(db['users'].keys())}")

# --- 中间件 ---
@app.before_request
def auth_middleware():
    if request.path.startswith('/static'): return
    db = load_auth_db()
    if not db["users"]:
        if request.endpoint != 'setup': return redirect(url_for('setup'))
        return
    if request.endpoint == 'setup': return redirect(url_for('login'))
    if request.endpoint in ['login', 'logout']: return
    session.permanent = True
    if not session.get('logged_in'): return redirect(url_for('login'))

# --- 路由 ---
@app.route('/setup', methods=['GET', 'POST'])
def setup():
    if request.method == 'POST':
        u, p = request.form.get('username'), request.form.get('password')
        if u and p:
            add_user_logic(u, p)
            return redirect(url_for('login'))
    return render_template('setup.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        u, p = request.form.get('username'), request.form.get('password')
        db = load_auth_db()
        if u in db["users"] and check_password_hash(db["users"][u], p):
            session['logged_in'] = True
            session['user'] = u
            return redirect(url_for('portal'))
        error = "用户名或密码错误"
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/')
def portal(): return render_template('portal.html')

# --- 拆分模块 ---
@app.route('/tool/splitter')
def splitter_ui(): return render_template('splitter.html')

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
        curr = 0
        for w in weights[:-1]:
            val = round((w / sum_weights) * total_qty, 1)
            if val == 0 and total_qty > 1: val = 0.1
            amounts.append(val)
            curr += val
        amounts.append(round(total_qty - curr, 1))
    return amounts

@app.route('/api/splitter/analyze', methods=['POST'])
def splitter_analyze():
    file = request.files.get('file')
    if not file: return jsonify({"error": "未找到文件"}), 400
    try:
        content = file.read()
        f_bytes = io.BytesIO(content)
        visible = []
        if file.filename.lower().endswith('.xlsx'):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(f_bytes, read_only=True)
                visible = [s.title for s in wb.worksheets if s.sheet_state == 'visible']
                wb.close()
            except: pass
        f_bytes.seek(0)
        xl = pd.ExcelFile(f_bytes)
        sheets = [s for s in xl.sheet_names if s in visible] if visible else xl.sheet_names
        if not sheets: sheets = xl.sheet_names
        df = xl.parse(sheets[0], nrows=10)
        return jsonify({"sheets": sheets, "columns": df.columns.tolist()})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/api/splitter/sheet_info', methods=['POST'])
def splitter_sheet_info():
    file = request.files.get('file')
    try:
        df = pd.read_excel(file, sheet_name=request.form.get('sheet_name'))
        unit_col = next((c for c in df.columns if '单位' in str(c)), None)
        units = df[unit_col].dropna().unique().tolist() if unit_col else []
        return jsonify({"columns": df.columns.tolist(), "units": [str(u) for u in units]})
    except Exception as e: return jsonify({"error": str(e)}), 500

@app.route('/api/splitter/process', methods=['POST'])
def splitter_process():
    # 获取参数：A, B, C 列
    col_a = request.form.get('col_a') # 数量 (必填)
    col_b = request.form.get('col_b') # 单价 (可选)
    col_c = request.form.get('col_c') # 金额 (可选)
    
    file = request.files.get('file')
    sheet = request.form.get('sheet_name')
    days = int(request.form.get('days', 12))
    selected_cols = request.form.getlist('cols')[:10]
    int_units = request.form.getlist('int_units')
    
    try:
        df = pd.read_excel(file, sheet_name=sheet)
        if col_a not in df.columns:
            return jsonify({"success": False, "error": f"未找到数量列: {col_a}"}), 400

        unit_col = next((c for c in df.columns if '单位' in str(c)), None)
        
        # 清洗数据
        df = df.dropna(subset=[col_a])
        df[col_a] = pd.to_numeric(df[col_a], errors='coerce')
        df = df[df[col_a] > 0]
        
        daily_rows = [[] for _ in range(days)]
        
        for _, row in df.iterrows():
            unit = str(row.get(unit_col, '')) if unit_col else ""
            orig_qty = row[col_a]
            if pd.isna(orig_qty) or orig_qty == 0: continue
            
            # 活跃天数逻辑
            if orig_qty <= 3: active = 1
            elif orig_qty <= 10: active = random.randint(2, min(4, days))
            else: active = random.randint(3, min(days, 10))
            
            splits = split_smart_algo(orig_qty, active, unit in int_units)
            indices = sorted(random.sample(range(days), len(splits)))
            
            for i, day_idx in enumerate(indices):
                new_qty = splits[i]
                ratio = new_qty / orig_qty
                
                new_row = {}
                for col in selected_cols:
                    if col not in row: continue
                    val = row[col]
                    
                    # --- 核心逻辑 A * B = C ---
                    
                    # 1. 数量列 (A) -> 直接替换
                    if col == col_a:
                        new_row[col] = new_qty
                        
                    # 2. 金额列 (C) -> 尝试 A * B 计算
                    elif col == col_c and col_b and (col_b in row):
                        try:
                            # 优先使用：新数量 * 原单价
                            price = float(row[col_b])
                            new_row[col] = round(new_qty * price, 2)
                        except:
                            # 如果单价无效，回退到按比例缩放原金额
                            try: new_row[col] = round(float(val) * ratio, 2)
                            except: new_row[col] = val
                            
                    # 3. 其他列 -> 如果是数字且看起来像总量（如重量），按比例缩放
                    elif isinstance(val, (int, float)):
                        # 排除 B 列（单价）和 ID 类列，防止被误改
                        is_static = (col == col_b) or any(k in str(col).lower() for k in ['id', 'code', 'date', '日期', '单价', '价', '规格'])
                        if not is_static:
                            try:
                                new_row[col] = round(float(val) * ratio, 2)
                            except:
                                new_row[col] = val
                        else:
                            new_row[col] = val
                    else:
                        new_row[col] = val
                
                daily_rows[day_idx].append(new_row)
        
        filename = f"拆分_{sheet}_{uuid.uuid4().hex[:8]}.xlsx"
        path = os.path.join(app.config['RESULT_FOLDER'], filename)
        
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            for i in range(days):
                pd.DataFrame(daily_rows[i]).to_excel(writer, sheet_name=f'第{i+1}天', index=False)
        
        return jsonify({"success": True, "filename": filename})
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/splitter/download/<filename>')
def splitter_download(filename):
    return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)

# --- 比对模块 (保持不变) ---
@app.route('/tool/compare')
def compare_ui(): return render_template('compare.html')

def clean_name_algo(text): return re.sub(r'\*.*?\*', '', str(text)).strip() if not pd.isna(text) else ""

@app.route('/api/compare/get_logs')
def compare_get_logs():
    global log_stream
    logs, log_stream = list(log_stream), []
    return jsonify(logs)

@app.route('/api/compare/get_headers', methods=['POST'])
def compare_get_headers():
    try: return jsonify({"columns": pd.read_excel(request.files['file'], nrows=1).columns.tolist()})
    except Exception as e: return jsonify({"error": str(e)})

@app.route('/api/compare/process', methods=['POST'])
def compare_process():
    try:
        f_in, f_out = request.files['file_in'], request.files['file_out']
        m = request.form
        df_in, df_out = pd.read_excel(f_in), pd.read_excel(f_out)
        df_in['__k'] = df_in[m['map_in_name']].apply(clean_name_algo)
        df_out['__k'] = df_out[m['map_out_name']].apply(clean_name_algo)
        agg_in = df_in.groupby('__k')[[m['map_in_qty'], m['map_in_val']]].sum().reset_index()
        agg_out = df_out.groupby('__k')[[m['map_out_qty'], m['map_out_val']]].sum().reset_index()
        agg_in.columns = ['Key', 'In_Qty', 'In_Val']
        agg_out.columns = ['Key', 'Out_Qty', 'Out_Val']
        res = pd.merge(agg_in, agg_out, on='Key', how='outer').fillna(0)
        res['Diff_Qty'] = res['Out_Qty'] - res['In_Qty']
        res['Diff_Val'] = res['Out_Val'] - res['In_Val']
        fname = f"result_{uuid.uuid4().hex}.xlsx"
        res.to_excel(os.path.join(app.config['RESULT_FOLDER'], fname), index=False)
        return jsonify({"success": True, "filename": fname})
    except Exception as e: return jsonify({"success": False, "message": str(e)})

@app.route('/api/compare/download/<filename>')
def compare_download(filename):
    return send_from_directory(app.config['RESULT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
