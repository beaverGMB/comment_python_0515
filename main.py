import tkinter as tk
import win32gui
import win32con
from screeninfo import get_monitors
from tkinterweb import HtmlFrame
import os
import threading
import queue
from flask import Flask, request, render_template_string, redirect, url_for, send_file, make_response
from flask_socketio import SocketIO, emit
import pandas as pd
import io
from datetime import datetime
import tkinter.filedialog as filedialog
from PIL import Image, ImageTk
from uuid import uuid4

# ==== Flask + SocketIO ====
app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*", async_mode="eventlet")

message_queue = queue.Queue()
message_log = []
messages = []
server_session_id = str(uuid4())

# ==== HTML テンプレート ====
TEMPLATE_HTML = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">  <!-- スマホ対応追加 -->
    <title>コメント送信フォーム</title>
    <script src="https://cdn.socket.io/4.5.4/socket.io.min.js"></script>

    <style>
        body {
            font-family: sans-serif;
            padding: 10px;
            margin: 0;
            font-size: 16px;
            background-color: #f9f9f9;
        }

        input, select, button {
            font-size: 1em;
            padding: 0.5em;
            margin: 0.3em 0;
            width: 100%;
            box-sizing: border-box;
        }

        form {
            margin-top: 10px;
        }

        h2 {
            font-size: 1.2em;
        }

        ul {
            padding-left: 1em;
            list-style-type: none;
        }

        li {
            margin-bottom: 0.5em;
            background: #fff;
            padding: 0.5em;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .template-select {
            margin-top: 10px;
        }
    </style>

    <script>
        const socket = io();

        socket.on("new_comment", function(data) {
            const commentList = document.querySelector("ul");
            const li = document.createElement("li");
            li.innerHTML = `<strong>${data.name}</strong>: ${data.text}`;
            commentList.appendChild(li);
        });

        function getCookie(name) {
            let match = document.cookie.match(new RegExp('(^| )' + name + '=([^;]+)'));
            if (match) return decodeURIComponent(match[2]);
            return null;
        }

        function setCookie(name, value, days) {
            let expires = "";
            if (days) {
                let d = new Date();
                d.setTime(d.getTime() + (days*24*60*60*1000));
                expires = "; expires=" + d.toUTCString();
            }
            document.cookie = name + "=" + encodeURIComponent(value) + expires + "; path=/";
        }

        function saveNameAndReload() {
            const nameInput = document.getElementById("nameInput").value.trim();
            if (nameInput !== "") {
                setCookie("username", nameInput, 7);
                setCookie("session_id", "{{ server_session_id }}", 7);
                location.reload();
            } else {
                alert("名前を入力してください。");
            }
        }

        function setTemplateText(value) {
            document.getElementById('msg').value = value;
        }

        function loadUsernameToForm() {
            const name = getCookie("username");
            const session = getCookie("session_id");
            if (!name || session !== "{{ server_session_id }}") {
                document.getElementById("nameEntry").style.display = "block";
                document.getElementById("commentSection").style.display = "none";
            } else {
                document.getElementById("nameEntry").style.display = "none";
                document.getElementById("commentSection").style.display = "block";
                document.getElementById("hiddenName").value = name;
            }
        }

        window.onload = loadUsernameToForm;
    </script>
</head>
<body>
    <div id="nameEntry" style="display:none;">
        <h2>名前を入力してください</h2>
        <input type="text" id="nameInput" placeholder="名前を入力">
        <button onclick="saveNameAndReload()">OK</button>
    </div>

    <div id="commentSection" style="display:none;">
        <h2>コメント送信フォーム</h2>

        <label for="templates">定型文から選ぶ：</label>
        <select id="templates" onchange="setTemplateText(this.value)">
            <option value="">-------- 定型文を選択 --------</option>
            <option value="おはようございます">おはようございます</option>
            <option value="こんにちは">こんにちは</option>
            <option value="寒いです">寒いです</option>
            <option value="頑張れー">頑張れー</option>
        </select>

        <form method="POST" action="/comment">
            <input type="hidden" name="name" id="hiddenName">
            <input type="text" name="msg" id="msg" required placeholder="コメントを入力">
            <button type="submit">送信</button>
        </form>

        <hr>
        <form action="/download" method="get"> 
            <label>形式を選択：</label>
            <select name="format">
                <option value="xlsx">Excel (.xlsx)</option>
                <option value="csv">CSV (.csv)</option>
            </select>
            <button type="submit">ダウンロード</button>
        </form>

        <hr>
        <h2>コメント履歴</h2>
        <ul>
            {% for msg in messages %}
                <li><strong>{{ msg.name }}</strong>: {{ msg.text }}</li>
            {% endfor %}
        </ul>
    </div>
</body>
</html>
"""


@socketio.on("connect")
def handle_connect():
    print("接続を確認")


@app.route("/")
def form():
    return render_template_string(TEMPLATE_HTML, messages=message_log, server_session_id=server_session_id)

@app.route("/comment", methods=["POST"])
def comment():
    msg = request.form.get("msg", "")
    name = request.form.get("name", "名無し")
    if msg and name:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = {"name": name, "text": msg, "time": now}
        message_queue.put(entry)
        message_log.append(entry)
        socketio.emit("new_comment", entry)
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

# ==== Flaskサーバをスレッドで起動 ====
def run_flask():
    socketio.run(app, host="0.0.0.0", port=5000)

# ==== Tkinter GUI ====
def set_always_on_top(hwnd):
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

def create_menu_window():
    menu_root = tk.Toplevel()
    menu_root.title("コントロールメニュー")
    menu_root.geometry("300x300+50+50")
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
            except Exception as e:
                print(f"保存エラー: {e}")

    def clear_comments():
        messages.clear()
        message_log.clear()

    tk.Button(menu_root, text="コメント履歴表示", command=lambda: print(message_log)).pack(pady=5)
    tk.Button(menu_root, text="コメント履歴クリア", command=clear_comments).pack(pady=5)
    tk.Button(menu_root, text="CSV形式で保存", command=lambda: export_file_dialog("csv")).pack(pady=5)
    tk.Button(menu_root, text="Excel形式で保存", command=lambda: export_file_dialog("xlsx")).pack(pady=5)
    tk.Button(menu_root, text="アプリを終了", command=lambda: os._exit(0)).pack(pady=10)

def main():
    threading.Thread(target=run_flask, daemon=True).start()

    root = tk.Tk()
    root.title("コメント表示")
    root.overrideredirect(True)
    screen = get_monitors()[0]
    w, h = screen.width // 4, screen.height
    x, y = screen.width - w, 0
    root.geometry(f"{w}x{h}+{x}+{y}")
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

    line_height = 40
    max_comments = h // line_height
    last_html = [""]

    def update_comments():
        try:
            while True:
                new_entry = message_queue.get_nowait()
                messages.append(new_entry)
        except queue.Empty:
            pass

        if len(messages) > max_comments:
            messages.pop(0)

        body_content = "\n".join(
            f'''
                <div class="message-block">
                    <div class="username">{msg.get("name", "名無し")}</div>
                    <div class="bubble">{msg["text"]}</div>
                </div>
            '''
            for msg in messages
        )
        full_html = bubble_html.replace("</body>", f"{body_content}\n</body>")
        if full_html != last_html[0]:
            html_frame.load_html(full_html)
            last_html[0] = full_html

        root.after(1000, update_comments)

    create_menu_window()
    update_comments()
    root.mainloop()

if __name__ == "__main__":
    main()
