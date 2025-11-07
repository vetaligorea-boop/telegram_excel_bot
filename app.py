import os
import requests
from flask import Flask, request, jsonify
from processor import format_pub_zero, run_combined_flow

# ============================
# CONFIG
# ============================
BOT_TOKEN = "8548367764:AAHLp9wmOQcHsMWWLrvAIxAr_TzVxfJU-gg"  # <-- tokenul tău real
API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"
FILE_API_URL = f"https://api.telegram.org/file/bot{BOT_TOKEN}"
ALLOWED_EXTS = [".xls", ".xlsx", ".xlsm"]

app = Flask(__name__)

USER_STATE = {}

# ============================
# HELPER FUNCTIONS
# ============================

def get_user_state(chat_id):
    key = str(chat_id)
    if key not in USER_STATE:
        USER_STATE[key] = {}
    return USER_STATE[key]

def ensure_user_dir(chat_id):
    base = "files"
    os.makedirs(base, exist_ok=True)
    user_dir = os.path.join(base, str(chat_id))
    os.makedirs(user_dir, exist_ok=True)
    return user_dir

def send_message(chat_id, text, reply_markup=None):
    payload = {"chat_id": chat_id, "text": text, "parse_mode": "HTML"}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    requests.post(f"{API_URL}/sendMessage", json=payload)

def send_document(chat_id, file_path, caption=None):
    if not os.path.isfile(file_path):
        send_message(chat_id, f"⚠️ Nu găsesc fișierul: {file_path}")
        return
    with open(file_path, "rb") as f:
        files = {"document": f}
        data = {"chat_id": chat_id}
        if caption:
            data["caption"] = caption
        requests.post(f"{API_URL}/sendDocument", data=data, files=files)

def get_file(file_id):
    resp = requests.get(f"{API_URL}/getFile", params={"file_id": file_id})
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"Nu pot obține file_path pentru {file_id}: {data}")
    return data["result"]

def main_keyboard():
    return {
        "inline_keyboard": [
            [
                {"text": "📂 Trimit IN", "callback_data": "SEND_IN"},
                {"text": "📂 Trimit PUB_Zero", "callback_data": "SEND_PUB_ZERO"},
            ],
            [{"text": "🚀 Procesează", "callback_data": "PROCESS"}],
        ]
    }

# ============================
# HANDLE DOCUMENT
# ============================

def handle_document(message):
    chat_id = message["chat"]["id"]
    state = get_user_state(chat_id)
    role = state.get("await")

    if role not in ("IN", "PUB_ZERO"):
        send_message(chat_id, "👉 Te rog apasă mai întâi butonul pentru IN sau PUB_Zero.", main_keyboard())
        return

    if "document" not in message:
        send_message(chat_id, "❌ Nu am primit niciun fișier.")
        return

    file_id = message["document"]["file_id"]
    file_name = message["document"]["file_name"]
    _, ext = os.path.splitext(file_name)
    ext = ext.lower()

    if ext not in ALLOWED_EXTS:
        send_message(chat_id, "⚠️ Accept doar fișiere .xls, .xlsx sau .xlsm.")
        return

    info = get_file(file_id)
    file_url = f"{FILE_API_URL}/{info['file_path']}"
    user_dir = ensure_user_dir(chat_id)
    local_path = os.path.join(user_dir, file_name)

    resp = requests.get(file_url)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(resp.content)

    if role == "IN":
        state["in_path"] = local_path
        send_message(chat_id, "✅ Am salvat fișierul IN.", main_keyboard())
    elif role == "PUB_ZERO":
        state["pub_zero_path"] = local_path
        send_message(chat_id, "✅ Am salvat fișierul PUB_Zero.", main_keyboard())

    state["await"] = None

    if state.get("in_path") and state.get("pub_zero_path"):
        send_message(chat_id, "✅ Ambele fișiere sunt pregătite!\nApasă „🚀 Procesează”.", main_keyboard())

# ============================
# HANDLE CALLBACK
# ============================

def handle_callback_query(callback):
    data = callback.get("data")
    message = callback.get("message", {})
    chat_id = message["chat"]["id"]
    state = get_user_state(chat_id)

    if data == "SEND_IN":
        state["await"] = "IN"
        send_message(chat_id, "Trimite acum fișierul pentru IN (.xls/.xlsx/.xlsm).")

    elif data == "SEND_PUB_ZERO":
        state["await"] = "PUB_ZERO"
        send_message(chat_id, "Trimite acum fișierul pentru PUB_Zero (.xls/.xlsx/.xlsm).")

    elif data == "PROCESS":
        in_path = state.get("in_path")
        pub_zero_path = state.get("pub_zero_path")

        missing = []
        if not in_path:
            missing.append("- fișierul IN")
        if not pub_zero_path:
            missing.append("- fișierul PUB_Zero")

        if missing:
            send_message(chat_id, "⚠️ Încă lipsesc fișiere:\n" + "\n".join(missing), main_keyboard())
            return

        send_message(chat_id, "⏳ Procesez fișierele...")

        try:
            pub_in_path = format_pub_zero(pub_zero_path)
            final_path = run_combined_flow(in_path, pub_in_path)

            send_document(chat_id, pub_in_path, "📄 PUB_IN (generat din PUB_Zero)")
            send_document(chat_id, final_path, "📄 Desfasurator FINAL (IN_modificat) ✅")

        except Exception as e:
            send_message(chat_id, f"❌ Eroare la procesare: {e}")

# ============================
# FLASK ROUTES
# ============================

@app.route("/", methods=["GET"])
def index():
    return "✅ Botul este online."

@app.route("/", methods=["POST"])
def webhook():
    update = request.get_json(force=True)
    if "message" in update:
        msg = update["message"]
        chat_id = msg["chat"]["id"]
        text = msg.get("text", "")

        if text == "/start":
            USER_STATE[str(chat_id)] = {}
            send_message(
                chat_id,
                "👋 Salut!\n\n"
                "1️⃣ Apasă „📂 Trimit IN” și trimite fișierul IN (.xls/.xlsx/.xlsm)\n"
                "2️⃣ Apasă „📂 Trimit PUB_Zero” și trimite fișierul PUB_Zero\n"
                "3️⃣ Apasă „🚀 Procesează” pentru a primi PUB_IN și FINAL.",
                main_keyboard(),
            )
        elif "document" in msg:
            handle_document(msg)
        else:
            send_message(chat_id, "Folosește butoanele de mai jos 👇", main_keyboard())

    if "callback_query" in update:
        handle_callback_query(update["callback_query"])

    return jsonify(ok=True)

# ============================
# RUN LOCALLY (DEBUG)
# ============================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
