# ========== 標準ライブラリ ==========
import os
import sys
import re
import io
import queue
import threading
import shutil
import subprocess
import atexit
from datetime import datetime
from uuid import uuid4

# ========== GUI / Windows 系 ==========
import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog as filedialog
import win32gui
import win32con
from screeninfo import get_monitors
from tkinterweb import HtmlFrame
from playsound import playsound

# ========== Web / データ ==========
from flask import Flask, request, render_template, redirect, url_for, send_file
from flask_socketio import SocketIO
import pandas as pd

# ------------------------------------------------------------
# パス解決 (PyInstaller ― _MEIPASS)
# ------------------------------------------------------------
BASE_DIR = getattr(sys, "_MEIPASS", os.path.abspath("."))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")
SOUND_PATH = os.path.join(BASE_DIR, "sounds", "samplesound.mp3")
BUBBLE_HTML_PATH = os.path.join(BASE_DIR, "bubble.html")

# ------------------------------------------------------------
# Flask アプリ
# ------------------------------------------------------------
app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="threading")

# ------------------------------------------------------------
# 共有状態
# ------------------------------------------------------------
message_queue: "queue.Queue[dict]" = queue.Queue()
message_log: list[dict] = []          # 保存用
messages: list[dict] = []             # 表示用
unsaved_changes = False
SERVER_SESSION_ID = str(uuid4())

# ------------------------------------------------------------
# Tailscale Funnel
# ------------------------------------------------------------
FUNNEL_PORT = "5050"

def _tailscale_exists() -> bool:
    return shutil.which("tailscale") is not None

def start_tailscale_funnel() -> None:
    if not _tailscale_exists():
        print("[WARN] tailscale コマンドが無いので Funnel は起動しません。")
        return
    subprocess.run(["tailscale", "funnel", "--bg", FUNNEL_PORT], check=False)
    print(f"[INFO] Funnel started on {FUNNEL_PORT}")

def stop_tailscale_funnel() -> None:
    if not _tailscale_exists():
        return
    subprocess.run(["tailscale", "funnel", "--bg", FUNNEL_PORT, "off"], check=False)
    print("[INFO] Funnel stopped")

atexit.register(stop_tailscale_funnel)

# ------------------------------------------------------------
# Flask ルーティング
# ------------------------------------------------------------
@app.route("/")
def form():
    return render_template(
        "web_index.html",
        messages=list(reversed(message_log)),  # 最新を上に
        server_session_id=SERVER_SESSION_ID,
    )

@app.route("/comment", methods=["POST"])
def comment():  # noqa: D401
    msg = request.form.get("msg", "")
    name = request.form.get("name", "名無し")
    real_name = request.form.get("real_name", "")

    # HTML タグ禁止
    if re.search(r"<[^>]+>", msg + name + real_name):
        return (
            "<script>alert('HTMLタグは禁止です');window.history.back();</script>",
            400,
        )

    if msg and name:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = {"name": name, "real_name": real_name, "text": msg, "time": now}
        message_queue.put(entry)
        message_log.append(entry)
        socketio.emit("new_comment", entry)

        global unsaved_changes
        unsaved_changes = True
        return redirect(url_for("form"))
    return "エラー", 400

@app.route("/download")
def download_file():
    if not message_log:
        return "データがありません", 404

    fmt = request.args.get("format", "xlsx").lower()
    df = pd.DataFrame(message_log).rename(
        columns={"real_name": "本名", "name": "名前", "text": "コメント", "time": "時刻"}
    )

    if fmt == "csv":
        buf = io.StringIO()
        df.to_csv(buf, index=False, encoding="utf-8-sig")
        return send_file(
            io.BytesIO(buf.getvalue().encode("utf-8-sig")),
            download_name="comments.csv",
            as_attachment=True,
            mimetype="text/csv",
        )
    else:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="コメント履歴")
        out.seek(0)
        return send_file(
            out,
            download_name="comments.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ------------------------------------------------------------
# Flask 実行スレッド
# ------------------------------------------------------------
def run_flask():
    try:
        socketio.run(
            app,
            host="127.0.0.1",
            port=int(FUNNEL_PORT),
            debug=False,
            use_reloader=False,
            allow_unsafe_werkzeug=True,
        )
    except Exception as e:
        tk.messagebox.showerror("Flask 起動失敗", str(e))

# ------------------------------------------------------------
# Tkinter 関数
# ------------------------------------------------------------
def set_always_on_top(hwnd):
    win32gui.SetWindowPos(
        hwnd,
        win32con.HWND_TOPMOST,
        0, 0, 0, 0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
    )

def play_notification_sound():
    try:
        playsound(SOUND_PATH)
    except Exception as e:
        print(f"[WARN] 音声再生失敗: {e}")

def create_menu_window(switch_display_callback, root):
    menu = tk.Toplevel()
    menu.title("コントロールメニュー")
    menu.geometry("350x350")
    menu.attributes("-topmost", True)

    def export_dialog(fmt: str):
        df = pd.DataFrame(message_log).rename(
            columns={"real_name": "本名", "name": "名前", "text": "コメント", "time": "時刻"}
        )
        filetypes = [("Excelファイル", "*.xlsx")] if fmt == "xlsx" else [("CSVファイル", "*.csv")]
        ext = ".xlsx" if fmt == "xlsx" else ".csv"
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=filetypes)
        if not path:
            return
        try:
            if fmt == "csv":
                df.to_csv(path, index=False, encoding="utf-8-sig")
            else:
                df.to_excel(path, index=False)
            global unsaved_changes
            unsaved_changes = False
        except Exception as e:
            tk.messagebox.showerror("保存失敗", str(e))

    def confirm_exit():
        if unsaved_changes and not tk.messagebox.askyesno(
            "確認", "保存していないコメントは失われます。終了しますか？"
        ):
            return
        stop_tailscale_funnel()
        root.destroy()

    tk.Button(menu, text="表示モニター切替", command=switch_display_callback).pack(pady=5)
    tk.Button(menu, text="CSV で保存", command=lambda: export_dialog("csv")).pack(pady=5)
    tk.Button(menu, text="Excel で保存", command=lambda: export_dialog("xlsx")).pack(pady=5)
    tk.Button(menu, text="アプリ終了", command=confirm_exit).pack(pady=10)

# ------------------------------------------------------------
# メイン処理
# ------------------------------------------------------------
def main():
    # Flask 起動
    threading.Thread(target=run_flask, daemon=True).start()
    start_tailscale_funnel()

    # Tkinter ウィンドウ
    root = tk.Tk()
    root.title("コメント表示")
    root.overrideredirect(True)

    monitors = get_monitors()
    current_monitor = [0]

    def update_monitor_position():
        scr = monitors[current_monitor[0]]
        w, h = scr.width // 4, scr.height
        x, y = scr.x + scr.width - w, scr.y
        root.geometry(f"{w}x{h}+{x}+{y}")

    update_monitor_position()
    root.configure(bg="#fefefe")
    root.attributes("-topmost", True)
    root.update()
    set_always_on_top(root.winfo_id())

    wrapper = tk.Frame(root, bg="#fefefe")
    wrapper.pack(expand=True, fill="both")

    html_frame = HtmlFrame(wrapper, horizontal_scrollbar="auto", vertical_scrollbar="auto")
    html_frame.pack(expand=True, fill="both")

    # bubble.html 読み込み
    with open(BUBBLE_HTML_PATH, encoding="utf-8") as fp:
        bubble_html = fp.read()
    last_html = [""]

    def update_comments():
        new_added = False
        try:
            while True:
                messages.append(message_queue.get_nowait())
                new_added = True
        except queue.Empty:
            pass

        if new_added:
            threading.Thread(target=play_notification_sound, daemon=True).start()

        body = "\n".join(
            f"""
            <div class="comment-wrapper">
              <div class="shadow-box"></div>
              <div class="comment-box">
                <div class="name-label">{m['name']}</div>
                <div class="comment-name-time">
                  <span>　</span>
                  <span style='font-weight:normal;color:#666;'>{m['time'][11:16]}</span>
                </div>
                <div class="comment-text">{m['text']}</div>
                <div class="like"></div>
              </div>
            </div>
            """
            for m in reversed(messages)   # 最新を上
        )
        full_html = bubble_html.replace("</body>", f"{body}</body>")

        if full_html != last_html[0]:
            html_frame.load_html(full_html)
            last_html[0] = full_html
            root.after(200, lambda: html_frame.yview_moveto(0.0))

        root.after(1000, update_comments)

    def switch_display():
        current_monitor[0] = (current_monitor[0] + 1) % len(monitors)
        update_monitor_position()

    create_menu_window(switch_display, root)
    update_comments()

    def on_close():
        stop_tailscale_funnel()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    main()