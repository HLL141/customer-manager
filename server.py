#!/usr/bin/env python3
"""客户管理后台 — Flask 版，支持 Render 免费部署 + 百度网盘集成"""
import os
import json
import urllib.parse
import urllib.request
import hashlib
import io
import time
import openpyxl
from flask import Flask, request, jsonify, send_from_directory, redirect

app = Flask(__name__, static_url_path='', static_folder=os.path.dirname(os.path.abspath(__file__)))

# ═══ 配置 ═══
PORT = int(os.environ.get('PORT', 8765))
# 云部署时用相对路径（当前目录），本地用 E 盘
if os.environ.get('RAILWAY_ENVIRONMENT') or os.environ.get('RENDER'):
    DATA_DIR = os.path.dirname(os.path.abspath(__file__))
else:
    DATA_DIR = r'E:\新建文件夹'
DATA_FILE = os.path.join(DATA_DIR, 'customers.json')
PUBLIC_URL = os.environ.get('PUBLIC_URL', '').rstrip('/')

# ═══ 百度网盘配置 ═══
BAIDU_APP_KEY = "vUNpOCsK4nFa1TO7ghAR5G6imZUrSrfv"
BAIDU_SECRET_KEY = "n6gAMoMyAEEnzPaa0L3KCnPDTBNdm21A"
BAIDU_CONFIG_FILE = os.path.join(DATA_DIR, 'baidu_config.json')
BAIDU_AUTH_URL = "https://openapi.baidu.com/oauth/2.0/authorize"
BAIDU_TOKEN_URL = "https://openapi.baidu.com/oauth/2.0/token"
BAIDU_API_PRECREATE = "https://pan.baidu.com/rest/2.0/xpan/file?method=precreate"
BAIDU_API_CREATE = "https://pan.baidu.com/rest/2.0/xpan/file?method=create"
BAIDU_UPLOAD_BASE = "https://d.pcs.baidu.com/rest/2.0/pcs/superfile2"


def get_redirect_uri():
    """OAuth 回调地址 — 云端固定地址"""
    return 'https://web-production-929b.up.railway.app/oauth/callback'


# ═══ 客户数据读写 ═══
def load_customers():
    try:
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
    except Exception as e:
        print(f"[WARN] 读取数据失败: {e}")
    return [
        {"name": "张三", "phone": "13800138001", "plan": "基础版", "note": "微信来源", "date": "2026-06-08"},
        {"name": "李四", "phone": "13900139002", "plan": "进阶版", "note": "老客户推荐", "date": "2026-06-05"},
    ]


def save_customers(data):
    os.makedirs(os.path.dirname(DATA_FILE), exist_ok=True)
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[OK] 已保存 {len(data)} 条记录 -> {DATA_FILE}")


# ═══ 百度网盘 Token 管理 ═══
def load_baidu_config():
    try:
        if os.path.exists(BAIDU_CONFIG_FILE):
            with open(BAIDU_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_baidu_config(config):
    os.makedirs(os.path.dirname(BAIDU_CONFIG_FILE), exist_ok=True)
    with open(BAIDU_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print("[OK] 百度网盘配置已保存")


def get_valid_token():
    config = load_baidu_config()
    token = config.get('access_token')
    refresh = config.get('refresh_token')
    expires = config.get('expires_at', 0)

    if token and time.time() < expires - 3600:
        return token

    if refresh:
        try:
            data = urllib.parse.urlencode({
                'grant_type': 'refresh_token',
                'refresh_token': refresh,
                'client_id': BAIDU_APP_KEY,
                'client_secret': BAIDU_SECRET_KEY,
            }).encode('utf-8')
            req = urllib.request.Request(BAIDU_TOKEN_URL, data=data)
            with urllib.request.urlopen(req, timeout=15) as resp:
                result = json.loads(resp.read().decode('utf-8'))
            if 'access_token' in result:
                config['access_token'] = result['access_token']
                config['refresh_token'] = result.get('refresh_token', refresh)
                config['expires_at'] = time.time() + result.get('expires_in', 2592000)
                save_baidu_config(config)
                print("[OK] Token 已刷新")
                return config['access_token']
        except Exception as e:
            print(f"[WARN] Token 刷新失败: {e}")
    return None


# ═══ 百度网盘文件上传 ═══
def upload_to_baidu(file_content, filename, custom_path=None):
    token = get_valid_token()
    if not token:
        return {"ok": False, "error": "未绑定百度网盘或 Token 已过期，请重新绑定"}

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    name, ext = os.path.splitext(filename)
    base_dir = custom_path.rstrip('/') if custom_path else '/apps/客户数据云备份'
    remote_path = f"{base_dir}/{name}_{timestamp}{ext}"
    file_size = len(file_content)
    content_md5 = hashlib.md5(file_content).hexdigest()

    # Step1: 预上传
    try:
        precreate_body = urllib.parse.urlencode({
            'path': remote_path, 'size': file_size, 'isdir': '0',
            'rtype': '3', 'autoinit': '1',
            'block_list': json.dumps([content_md5]),
        }).encode('utf-8')
        req = urllib.request.Request(f"{BAIDU_API_PRECREATE}&access_token={token}", data=precreate_body)
        with urllib.request.urlopen(req, timeout=30) as resp:
            pre_result = json.loads(resp.read().decode('utf-8'))
        uploadid = pre_result.get('uploadid')
        if not uploadid:
            return {"ok": False, "error": f"预上传返回异常: {pre_result}"}
    except Exception as e:
        return {"ok": False, "error": f"预上传失败: {e}"}

    # Step2: 上传文件内容
    try:
        boundary = '----WebKitFormBoundary7MA4YWxkTrZu0gW'
        body = (
            f'--{boundary}\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/octet-stream\r\n\r\n'
        ).encode('utf-8') + file_content + f'\r\n--{boundary}--\r\n'.encode('utf-8')

        upload_url = (
            f"{BAIDU_UPLOAD_BASE}"
            f"?method=upload&access_token={token}&type=tmpfile"
            f"&path={urllib.parse.quote(remote_path, safe='')}"
            f"&uploadid={uploadid}&partseq=0"
        )
        req = urllib.request.Request(upload_url, data=body)
        req.add_header('Content-Type', f'multipart/form-data; boundary={boundary}')
        with urllib.request.urlopen(req, timeout=60) as resp:
            up_result = json.loads(resp.read().decode('utf-8'))
    except Exception as e:
        return {"ok": False, "error": f"文件上传失败: {e}"}

    # Step3: 创建文件
    try:
        create_body = urllib.parse.urlencode({
            'path': remote_path, 'size': file_size, 'isdir': '0',
            'uploadid': uploadid,
            'block_list': json.dumps([content_md5]), 'rtype': '3',
        }).encode('utf-8')
        req = urllib.request.Request(f"{BAIDU_API_CREATE}&access_token={token}", data=create_body)
        with urllib.request.urlopen(req, timeout=30) as resp:
            create_result = json.loads(resp.read().decode('utf-8'))
    except Exception as e:
        return {"ok": False, "error": f"文件创建失败: {e}"}

    return {"ok": True, "path": remote_path, "size": file_size,
            "fs_id": create_result.get('fs_id', 0)}


# ═══ 路由 ═══
@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'customer.html')

@app.route('/customer.html')
def customer_html():
    return send_from_directory(app.static_folder, 'customer.html')

@app.route('/api/customers', methods=['GET'])
def api_get_customers():
    return jsonify(load_customers())

@app.route('/api/customers', methods=['POST'])
def api_post_customers():
    try:
        data = request.get_json()
        if isinstance(data, list):
            save_customers(data)
            return jsonify({"ok": True, "count": len(data)})
        return jsonify({"ok": False, "error": "数据格式应为数组"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route('/api/baidu/auth')
def baidu_auth():
    redirect_uri = get_redirect_uri()
    params = urllib.parse.urlencode({
        'response_type': 'code', 'client_id': BAIDU_APP_KEY,
        'redirect_uri': redirect_uri, 'scope': 'basic,netdisk', 'display': 'page',
    })
    return redirect(f"{BAIDU_AUTH_URL}?{params}")

@app.route('/oauth/callback')
def baidu_callback():
    code = request.args.get('code')
    error = request.args.get('error')
    redirect_uri = get_redirect_uri()

    if error:
        msg = urllib.parse.unquote(request.args.get('error_description', error))
        return redirect(f'/?auth_error={urllib.parse.quote(msg)}')

    if not code:
        return redirect('/?auth_error=未收到授权码')

    try:
        data = urllib.parse.urlencode({
            'grant_type': 'authorization_code', 'code': code,
            'client_id': BAIDU_APP_KEY, 'client_secret': BAIDU_SECRET_KEY,
            'redirect_uri': redirect_uri,
        }).encode('utf-8')
        req = urllib.request.Request(BAIDU_TOKEN_URL, data=data)
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode('utf-8'))

        if 'access_token' in result:
            config = {
                'access_token': result['access_token'],
                'refresh_token': result.get('refresh_token', ''),
                'expires_at': time.time() + result.get('expires_in', 2592000),
                'bound_at': time.strftime('%Y-%m-%d %H:%M:%S'),
            }
            save_baidu_config(config)
            return redirect('/?auth_success=1')
        else:
            err = result.get('error_description', result.get('error', '未知错误'))
            return redirect(f'/?auth_error={urllib.parse.quote(str(err))}')
    except Exception as e:
        return redirect(f'/?auth_error={urllib.parse.quote(str(e))}')

@app.route('/api/baidu/status')
def baidu_status():
    config = load_baidu_config()
    token = get_valid_token()
    return jsonify({
        "bound": bool(token and config.get('access_token')),
        "bound_at": config.get('bound_at', '')
    })

@app.route('/api/baidu/upload', methods=['POST'])
def baidu_upload():
    try:
        body = request.get_json()
        # 兼容新旧格式：旧格式直接是数组，新格式是 {customers: [...], path: "..."}
        if isinstance(body, list):
            customers = body
            custom_path = None
        else:
            customers = body.get('customers', [])
            custom_path = body.get('path', None)

        if not isinstance(customers, list):
            return jsonify({"ok": False, "error": "数据格式错误"}), 400

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "客户列表"
        ws.append(["姓名", "手机号码", "套餐", "备注", "添加日期"])
        for c in customers:
            ws.append([c.get('name', ''), c.get('phone', ''), c.get('plan', ''),
                       c.get('note', ''), c.get('date', '')])
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 16
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 24
        ws.column_dimensions['E'].width = 14

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        result = upload_to_baidu(buf.read(), "客户数据.xlsx", custom_path)
        return jsonify(result), 200 if result.get('ok') else 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route('/api/baidu/unbind', methods=['POST'])
def baidu_unbind():
    try:
        if os.path.exists(BAIDU_CONFIG_FILE):
            os.remove(BAIDU_CONFIG_FILE)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == '__main__':
    print(f"启动端口: {PORT}")
    print(f"公网地址: {PUBLIC_URL or '本地模式'}")
    app.run(host='0.0.0.0', port=PORT, debug=True)
