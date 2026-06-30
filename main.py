"""建筑规范查询对比工具 v2 - 云端更新+离线缓存"""

import os
import sys
import json
import time
import hashlib
import webbrowser
import threading
import urllib.request
import urllib.error
from flask import Flask, request, jsonify, session, render_template
from spec_lib import SpecLib

# ── 配置 ──
ADMIN_PASSWORD = "guifan2024"
# 云端更新地址（管理员部署后修改此URL）
REMOTE_BASE = "https://gitee.com/YOUR_USERNAME/spec-lib/raw/main"
REMOTE_EXCEL_URL = REMOTE_BASE + "/规范库.xlsx"
REMOTE_VERSION_URL = REMOTE_BASE + "/version.json"


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_internal_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
INTERNAL_DIR = get_internal_dir()
DATA_DIR = os.path.join(APP_DIR, "data")
LOCAL_EXCEL = os.path.join(DATA_DIR, "规范库.xlsx")
LOCAL_VERSION = os.path.join(DATA_DIR, "version.json")

os.makedirs(DATA_DIR, exist_ok=True)


def file_md5(path):
    """计算文件MD5"""
    if not os.path.exists(path):
        return ""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def get_local_version():
    """读取本地版本信息"""
    if os.path.exists(LOCAL_VERSION):
        try:
            with open(LOCAL_VERSION, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"version": "0", "md5": "", "update_time": ""}


def save_local_version(info):
    """保存本地版本信息"""
    with open(LOCAL_VERSION, "w", encoding="utf-8") as f:
        json.dump(info, f, ensure_ascii=False, indent=2)


def check_and_download_update():
    """检查云端更新，有新版本则下载"""
    try:
        req = urllib.request.Request(REMOTE_VERSION_URL, headers={"User-Agent": "SpecLibTool/1.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            remote_info = json.loads(resp.read().decode("utf-8"))

        local_info = get_local_version()

        # 版本号比较
        if remote_info.get("version", "0") <= local_info.get("version", "0"):
            return False, "已是最新版本"

        # 下载新Excel
        excel_req = urllib.request.Request(REMOTE_EXCEL_URL, headers={"User-Agent": "SpecLibTool/1.0"})
        with urllib.request.urlopen(excel_req, timeout=30) as resp:
            data = resp.read()

        # 保存到本地
        with open(LOCAL_EXCEL, "wb") as f:
            f.write(data)

        # 更新本地版本信息
        remote_info["md5"] = file_md5(LOCAL_EXCEL)
        remote_info["update_time"] = time.strftime("%Y-%m-%d %H:%M")
        save_local_version(remote_info)

        return True, f"已更新到 v{remote_info['version']}"

    except urllib.error.URLError:
        return False, "无法连接服务器（离线模式）"
    except Exception as e:
        return False, f"更新检查失败：{str(e)}"


# 加载规范库：优先本地缓存，其次exe同目录，最后内置
def load_spec_lib():
    if os.path.exists(LOCAL_EXCEL):
        return SpecLib(LOCAL_EXCEL)
    fallback = os.path.join(APP_DIR, "规范库.xlsx")
    if os.path.exists(fallback):
        return SpecLib(fallback)
    return SpecLib(None)


lib = load_spec_lib()
print(f"[启动] 规范库加载：{lib.filename}，共 {lib.record_count} 条记录")
print(f"[启动] APP_DIR={APP_DIR}, LOCAL_EXCEL={LOCAL_EXCEL}")

app = Flask(__name__, template_folder=os.path.join(INTERNAL_DIR, "templates"))
app.secret_key = "spec_lib_tool_2024"
# 将lib存储到app.config，确保路由中始终可访问
app.config["SPEC_LIB"] = lib


def get_lib():
    """获取规范库对象，优先从app.config获取"""
    from flask import current_app
    return current_app.config.get("SPEC_LIB") or lib


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/database")
def database_page():
    return render_template("database.html")


@app.route("/api/database")
def api_database():
    current_lib = get_lib()
    records = []
    for i, r in enumerate(current_lib.records):
        records.append({
            "id": i + 1,
            "category": r.get("category", ""),
            "name": r["name"],
            "code": r["code"],
            "status": r["status"] if r["status"] else "未标注",
            "impl_date": r["impl_date"],
            "strike": r.get("strike", False),
        })
    return jsonify({"records": records, "total": len(records)})


@app.route("/api/status")
def api_status():
    current_lib = get_lib()
    status = current_lib.get_status()
    status["_debug_record_count"] = current_lib.record_count
    status["_debug_filename"] = current_lib.filename
    local_info = get_local_version()
    status["data_version"] = local_info.get("version", "未知")
    status["data_source"] = "云端缓存" if os.path.exists(LOCAL_EXCEL) else "本地文件"
    return jsonify(status)


@app.route("/api/parse", methods=["POST"])
def api_parse():
    text = request.json.get("text", "")
    return jsonify(get_lib().parse_input(text))


@app.route("/api/compare", methods=["POST"])
def api_compare():
    data = request.json
    current_lib = get_lib()
    if "items" in data:
        results = current_lib.compare(data["items"])
    else:
        text = data.get("text", "")
        parsed = current_lib.parse_input(text)
        results = current_lib.compare(parsed)
    return jsonify(results)


@app.route("/api/login", methods=["POST"])
def api_login():
    pwd = request.json.get("password", "")
    if pwd == ADMIN_PASSWORD:
        session["admin"] = True
        return jsonify({"success": True})
    return jsonify({"success": False, "message": "密码错误"})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    if not session.get("admin"):
        return jsonify({"success": False, "message": "请先输入管理员密码"})
    if "file" not in request.files:
        return jsonify({"success": False, "message": "未选择文件"})
    f = request.files["file"]
    if not f.filename.endswith((".xlsx", ".xls")):
        return jsonify({"success": False, "message": "请选择Excel文件"})
    save_path = os.path.join(DATA_DIR, "规范库.xlsx")
    try:
        f.save(save_path)
        current_lib = get_lib()
        current_lib.load(save_path)
        # 同步更新app.config中的引用
        app.config["SPEC_LIB"] = current_lib
        # 更新本地版本
        local_info = get_local_version()
        local_info["version"] = str(int(local_info.get("version", "0")) + 1)
        local_info["md5"] = file_md5(save_path)
        local_info["update_time"] = time.strftime("%Y-%m-%d %H:%M")
        save_local_version(local_info)
        status = current_lib.get_status()
        return jsonify({"success": True, "message": f"规范库已更新：{status['filename']}，共{status['record_count']}条记录"})
    except Exception as e:
        return jsonify({"success": False, "message": f"加载失败：{str(e)}"})


@app.route("/api/check_update")
def api_check_update():
    """手动检查云端更新"""
    updated, msg = check_and_download_update()
    current_lib = get_lib()
    if updated:
        current_lib.load(LOCAL_EXCEL)
        app.config["SPEC_LIB"] = current_lib
    return jsonify({"updated": updated, "message": msg, "status": current_lib.get_status()})


def open_browser():
    webbrowser.open("http://127.0.0.1:5000")


def startup_update():
    """启动时后台检查更新"""
    updated, msg = check_and_download_update()
    if updated:
        current_lib = get_lib()
        current_lib.load(LOCAL_EXCEL)
        app.config["SPEC_LIB"] = current_lib
    print(f"[更新检查] {msg}")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # 启动时后台检查更新（不阻塞启动）
    threading.Thread(target=startup_update, daemon=True).start()
    if port == 5000:
        threading.Timer(1.0, open_browser).start()
        print("建筑规范查询对比工具已启动，浏览器将自动打开...")
        print("访问地址：http://127.0.0.1:5000")
        print("关闭此窗口即可停止服务")
    else:
        print(f"建筑规范查询对比工具已启动，端口：{port}")
    app.run(host="0.0.0.0", port=port, debug=False)
