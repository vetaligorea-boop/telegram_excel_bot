import os
import requests
from flask import Flask, request, jsonify
from processor import format_pub_zero, run_combined_flow

app = Flask(__name__)

# ================== CONFIG ==================

# Tokenul botului tau din BotFather.
# In Render trebuie setat ca Environment Variable cu numele BOT_TOKEN.
BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Lipseste BOT_TOKEN. Seteaza variabila BOT_TOKEN in Render.")

TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

# chat_id -> stare:
# "await": None / "IN" / "PUB_ZERO"
# "in_path": path fisier IN pe server
# "pub_zero_path": path fisier PUB_Zero pe server
USER_STATE = {}


# ================== HELPERI ==================

def get_user_dir(chat_id: int) -> str:
    """
    Folder separat pentru fiecare utilizator (chat).
    Fisierele se salveaza in /tmp, nu pe discul tau.
    """
    base = "/tmp/telegram_excel_bot"
    path = os.path.join(base, str(chat_id))
    os.makedirs(path, exist_ok=True)
    return path


def send_message(chat_id, text, keyboard=False):
    payload = {
        "chat_id": chat_id,
        "text": text,
    }
    if keyboard:
        payload["reply_markup"] = {
            "keyboard": [
                [{"text": "📂 Trimit IN"}, {"text": "📂 Trimit PUB_Zero"}],
                [{"text": "🚀 Proceseaza"}],
            ],
            "resize_keyboard": True,
            "one_time_keyboard": False,
        }
    try:
        requests.post(f"{TELEGRAM_API}/sendMessage", json=payload, timeout=10)
    except Exception:
        # aici am putea loga eroarea daca vrem
        pass


def download_file(file_id: str, dest_path: str) -> bool:
    """
    Descarca fisierul trimis pe Telegram in dest_path.
    """
    # 1) luam file_path de la Telegram
    try:
        r = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id}, timeout=10)
        data = r.json()
    except Exception:
        return False

    if not data.get("ok"):
        return False

    file_path = data["result"]["file_path"]
    url = f"{TELEGRAM_FILE_API}/{file_path}"

    # 2) descarcam continutul
    try:
        resp = requests.get(url, timeout=30)
    except Exception:
        return False

    if resp.status_code != 200:
        return False

    try:
        with open(dest_path, "wb") as f:
            f.write(resp.content)
    except Exception:
        return False

    return True


def send_file(chat_id: int, file_path: str, caption: str = ""):
    """
    Trimite un fisier inapoi utilizatorului prin Telegram.
    """
    if not os.path.isfile(file_path):
        send_message(chat_id, f"Nu gasesc fisierul: {os.path.basename(file_path)}")
        return

    with open(file_path, "rb") as f:
        files = {"document": (os.path.basename(file_path), f)}
        data = {"chat_id": chat_id, "caption": caption}
        try:
            requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files, timeout=60)
        except Exception:
            # daca pica, trimitem macar un mesaj text
            send_message(chat_id, "Nu am reusit sa trimit fisierul (eroare retea).")


# ================== HEALTHCHECK ==================

@app.route("/", methods=["GET"])
def index():
    return "Bot online ✅", 200


# ================== WEBHOOK ==================

@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}

    message = update.get("message")
    if not message:
        return jsonify(ok=True)

    chat = message.get("chat") or {}
    chat_id = chat.get("id")
    if not chat_id:
        return jsonify(ok=True)

    # init stare pentru acest chat
    if chat_id not in USER_STATE:
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}

    # 1) daca a trimis document
    if "document" in message:
        return handle_document(chat_id, message["document"])

    # 2) daca a trimis text
    text = (message.get("text") or "").strip()

    # /start – reset + instructiuni
    if text.startswith("/start"):
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}
        send_message(
            chat_id,
            "Salut 👋\n\n"
            "Cum functionez:\n"
            "1️⃣ Apasa „📂 Trimit IN” si trimite fisierul IN (playlist).\n"
            "2️⃣ Apasa „📂 Trimit PUB_Zero” si trimite fisierul PUB_Zero.\n"
            "3️⃣ Apasa „🚀 Proceseaza” ca sa primesti fisierele modificate (_modificat).",
            keyboard=True,
        )
        return jsonify(ok=True)

    # a cerut sa trimita IN
    if text == "📂 Trimit IN":
        USER_STATE[chat_id]["await"] = "IN"
        send_message(chat_id, "Trimite acum fisierul pentru IN (Excel).")
        return jsonify(ok=True)

    # a cerut sa trimita PUB_Zero
    if text == "📂 Trimit PUB_Zero":
        USER_STATE[chat_id]["await"] = "PUB_ZERO"
        send_message(chat_id, "Trimite acum fisierul pentru PUB_Zero (Excel).")
        return jsonify(ok=True)

    # a cerut procesarea
    if text == "🚀 Proceseaza":
        return handle_process(chat_id)

    # orice alt text → re-explicam
    send_message(
        chat_id,
        "Te rog foloseste butoanele:\n"
        "📂 Trimit IN → apoi trimite fisierul IN\n"
        "📂 Trimit PUB_Zero → apoi trimite fisierul PUB_Zero\n"
        "🚀 Proceseaza → pentru a primi fisierele modificate.",
        keyboard=True,
    )
    return jsonify(ok=True)


# ================== HANDLE DOCUMENT ==================

def handle_document(chat_id: int, document: dict):
    state = USER_STATE.get(chat_id) or {"await": None, "in_path": None, "pub_zero_path": None}
    role = state.get("await")

    if role not in ("IN", "PUB_ZERO"):
        send_message(
            chat_id,
            "Mai intai apasa „📂 Trimit IN” sau „📂 Trimit PUB_Zero”, apoi trimite fisierul potrivit.",
            keyboard=True,
        )
        return jsonify(ok=True)

    file_id = document["file_id"]
    file_name = document.get("file_name") or "fisier.xlsx"

    user_dir = get_user_dir(chat_id)
    _, ext = os.path.splitext(file_name)
    if not ext:
        ext = ".xlsx"

    if role == "IN":
        dest_path = os.path.join(user_dir, "IN" + ext)
    else:
        dest_path = os.path.join(user_dir, "PUB_Zero" + ext)

    ok = download_file(file_id, dest_path)
    if not ok:
        send_message(chat_id, "Nu am reusit sa descarc fisierul. Incearca din nou.")
        return jsonify(ok=True)

    if role == "IN":
        state["in_path"] = dest_path
        send_message(chat_id, "Am salvat fisierul IN ✅")
    else:
        state["pub_zero_path"] = dest_path
        send_message(chat_id, "Am salvat fisierul PUB_Zero ✅")

    state["await"] = None
    USER_STATE[chat_id] = state

    # daca avem ambele fisiere, anuntam
    if state.get("in_path") and state.get("pub_zero_path"):
        send_message(chat_id, "Ambele fisiere sunt pregatite ✅\nApasa „🚀 Proceseaza”.", keyboard=True)

    return jsonify(ok=True)


# ================== HANDLE PROCESS ==================

def handle_process(chat_id: int):
    state = USER_STATE.get(chat_id) or {}
    in_path = state.get("in_path")
    pub_zero_path = state.get("pub_zero_path")

    if not in_path or not pub_zero_path:
        msg = "Nu pot porni procesarea. Lipsesc:\n"
        if not in_path:
            msg += "- fisierul IN\n"
        if not pub_zero_path:
            msg += "- fisierul PUB_Zero\n"
        msg += "\nFoloseste butoanele pentru a trimite fisierele."
        send_message(chat_id, msg, keyboard=True)
        return jsonify(ok=True)

    send_message(chat_id, "Procesez fisierele... ⏳")

    try:
        # 1) PUB_Zero -> PUB_IN (_modificat)
        pub_in_path = format_pub_zero(pub_zero_path)

        # 2) IN + PUB_IN -> IN_modificat
        final_path = run_combined_flow(in_path, pub_in_path)

        # trimitem ambele fisiere inapoi
        send_file(chat_id, pub_in_path, "PUB_IN (din PUB_Zero)")
        send_file(chat_id, final_path, "IN_modificat (din IN + PUB_IN)")

        send_message(chat_id, "Gata ✅ Ti-am trimis fisierele modificate.")
        # reset pentru urmatorul set
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}

    except Exception as e:
        send_message(chat_id, f"A aparut o eroare la procesare:\n{e}", keyboard=True)

    return jsonify(ok=True)


# ================== MAIN ==================

if __name__ == "__main__":
    # pentru rulare locala (Render seteaza PORT automat)
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
