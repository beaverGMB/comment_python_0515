import tkinter as tk
import tkinter.messagebox
import win32gui
import win32con
from screeninfo import get_monitors
from tkinterweb import HtmlFrame
import os
#from playsound import playsound
import threading
import queue
from flask import Flask, request, render_template, redirect, url_for, send_file
from flask_socketio import SocketIO
import pandas as pd
import io
from datetime import datetime
import tkinter.filedialog as filedialog
from uuid import uuid4
import re

# ==== Flask + SocketIO ====
app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="eventlet")

message_queue = queue.Queue()
message_log = []
messages = []
server_session_id = str(uuid4())

unsaved_changes = False  # ファイルを保存したのか否か

@app.route("/")
def form():
    return render_template("web_index.html", messages=message_log, server_session_id=server_session_id)

@app.route("/comment", methods=["POST"])
def comment():
    msg = request.form.get("msg", "")
    name = request.form.get("name", "名無し")

    html_tag_pattern = re.compile(r"<[^>]+>")
    if html_tag_pattern.search(msg) or html_tag_pattern.search(name):
        return '''
        <script>
            alert("HTMLタグは禁止されています。");
            window.history.back();
        </script>
        ''', 400

    if msg and name:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = {"name": name, "text": msg, "time": now}
        message_queue.put(entry)
        message_log.append(entry)
        socketio.emit("new_comment", entry)
        global unsaved_changes
        unsaved_changes = True #新規メッセージがあるたびにTrueにし、警告を出す
        return redirect(url_for("form"))
    return "エラー", 400

@app.route("/download")
def download_file():
    if not message_log:
        return "データがありません", 404

    file_format = request.args.get("format", "xlsx").lower()
    df = pd.DataFrame(message_log)
    df.rename(columns={"name": "名前", "text": "コメント", "time": "時刻"}, inplace=True)
    output = io.BytesIO()

    if file_format == "csv":
        output_text = io.StringIO()
        df.to_csv(output_text, index=False, encoding="utf-8-sig")
        output_text.seek(0)
        return send_file(
            io.BytesIO(output_text.read().encode("utf-8-sig")),
            download_name="comments.csv",
            as_attachment=True,
            mimetype="text/csv"
        )
    else:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="コメント履歴")
        output.seek(0)
        return send_file(
            output,
            download_name="comments.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def run_flask():
    socketio.run(app, host="0.0.0.0", port=5000)

def set_always_on_top(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def create_menu_window(switch_display_callback, root):
    menu_root = tk.Toplevel()
    menu_root.title("コントロールメニュー")
    menu_root.geometry("350x350")
    menu_root.attributes("-topmost", True)

    def export_file_dialog(format_type):
        df = pd.DataFrame(message_log)
        df.rename(columns={"name": "名前", "text": "コメント", "time": "時刻"}, inplace=True)
        filetypes = [("Excelファイル", "*.xlsx")] if format_type == "xlsx" else [("CSVファイル", "*.csv")]
        def_ext = ".xlsx" if format_type == "xlsx" else ".csv"
        filepath = filedialog.asksaveasfilename(defaultextension=def_ext, filetypes=filetypes)
        if filepath:
            try:
                if format_type == "csv":
                    df.to_csv(filepath, index=False, encoding="utf-8-sig")
                else:
                    df.to_excel(filepath, index=False)
                global unsaved_changes
                unsaved_changes = False
            except Exception as e:
                print(f"保存エラー: {e}")

    def confirm_exit():
        if unsaved_changes:
            confirm = tk.messagebox.askyesno("確認", "保存していない履歴は削除されてしまいます。よろしいですか？")
            if not confirm:
                return
        root.destroy()

    tk.Button(menu_root, text="表示モニター切り替え", command=switch_display_callback).pack(pady=5)
    tk.Button(menu_root, text="CSV形式で保存", command=lambda: export_file_dialog("csv")).pack(pady=5)
    tk.Button(menu_root, text="Excel形式で保存", command=lambda: export_file_dialog("xlsx")).pack(pady=5)
    tk.Button(menu_root, text="アプリを終了", command=confirm_exit).pack(pady=10)

def main():
    threading.Thread(target=run_flask, daemon=True).start()

    root = tk.Tk()
    root.title("コメント表示")
    root.overrideredirect(True)

    monitors = get_monitors()
    current_monitor = [0]

    def update_monitor_position():
        screen = monitors[current_monitor[0]]
        w, h = screen.width // 4, screen.height
        x, y = screen.x + screen.width - w, screen.y
        root.geometry(f"{w}x{h}+{x}+{y}")

    update_monitor_position()
    root.configure(bg="#fefefe")
    root.attributes("-topmost", True)
    root.update()
    hwnd = root.winfo_id()
    set_always_on_top(hwnd)

    wrapper = tk.Frame(root, bg="#fefefe")
    wrapper.pack(expand=True, fill="both")

    html_frame = HtmlFrame(wrapper, horizontal_scrollbar="auto", vertical_scrollbar="auto")
    html_frame.pack(expand=True, fill="both")

    with open("bubble.html", encoding="utf-8") as f:
        bubble_html = f.read()

    last_html = [""]

    def update_comments():
        try:
            while True:
                new_entry = message_queue.get_nowait()
                messages.append(new_entry)
        except queue.Empty:
            pass

        body_content = "\n".join(
            f'''
                <div class="comment-wrapper">
                    <div class="shadow-box"></div>
                        <div class="comment-box">
                        <div class="name-label">{msg.get("name")}</div>
                        <div class="comment-name-time">
                            <span>　</span>
                            <span style="font-weight: normal; color: #666;">{msg.get("time", "")[11:16]}</span>
                        </div>
                    <div class="comment-text">{msg["text"]}</div>
                        <div class="like"><!--ここにリアクションを表示--></div>
                    </div>
                </div>
            '''
            for msg in messages
        )
        full_html = bubble_html.replace("</body>", f"{body_content}</body>")

        if full_html != last_html[0]:
            html_frame.load_html(full_html)
            last_html[0] = full_html

            def scroll_to_bottom():
                html_frame.yview_moveto(1.0)

            root.after(200, scroll_to_bottom)

        root.after(1000, update_comments)

    def switch_display():
        current_monitor[0] = (current_monitor[0] + 1) % len(monitors)
        update_monitor_position()

    create_menu_window(switch_display, root)
    update_comments()
    root.mainloop()

if __name__ == "__main__":
    main()
