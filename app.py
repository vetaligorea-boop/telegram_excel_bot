import os
import requests
from flask import Flask, request, jsonify
from processor import format_pub_zero, run_combined_flow

app = Flask(__name__)

# ================== CONFIG ==================

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Lipseste BOT_TOKEN. Seteaza variabila BOT_TOKEN in Render.")

TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

# In memorie tinem DOAR ce fisier asteptam (IN sau PUB_ZERO).
# Fisierele reale le luam de pe disc, ca sa nu depindem de restarturi.
USER_STATE = {}  # chat_id -> {"await": None / "IN" / "PUB_ZERO"}


# ================== HELPERI ==================

def get_user_dir(chat_id: int) -> str:
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
        pass


def download_file(file_id: str, dest_path: str) -> bool:
    try:
        r = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id}, timeout=10)
        data = r.json()
    except Exception:
        return False

    if not data.get("ok"):
        return False

    file_path = data["result"]["file_path"]
    url = f"{TELEGRAM_FILE_API}/{file_path}"

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
    if not os.path.isfile(file_path):
        send_message(chat_id, f"Nu gasesc fisierul: {os.path.basename(file_path)}")
        return

    with open(file_path, "rb") as f:
        files = {"document": (os.path.basename(file_path), f)}
        data = {"chat_id": chat_id, "caption": caption}
        try:
            requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files, timeout=60)
        except Exception:
            send_message(chat_id, "Nu am reusit sa trimit fisierul (eroare retea).")


def find_latest_with_prefix(user_dir: str, prefix: str):
    """
    Cauta cel mai recent fisier din user_dir care incepe cu prefix.
    Ex: 'IN', 'PUB_Zero'.
    """
    if not os.path.isdir(user_dir):
        return None

    candidates = []
    for name in os.listdir(user_dir):
        if name.startswith(prefix):
            full = os.path.join(user_dir, name)
            if os.path.isfile(full):
                mtime = os.path.getmtime(full)
                candidates.append((mtime, full))

    if not candidates:
        return None

    # luam fisierul cel mai nou
    candidates.sort(reverse=True)
    return candidates[0][1]


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

    # init stare
    if chat_id not in USER_STATE:
        USER_STATE[chat_id] = {"await": None}

    # DOCUMENT?
    if "document" in message:
        return handle_document(chat_id, message["document"])

    # TEXT?
    text = (message.get("text") or "").strip()

    # /start
    if text.startswith("/start"):
        USER_STATE[chat_id] = {"await": None}
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

    # Astept fisier IN
    if text == "📂 Trimit IN":
        USER_STATE[chat_id]["await"] = "IN"
        send_message(chat_id, "Trimite acum fisierul pentru IN (Excel).")
        return jsonify(ok=True)

    # Astept fisier PUB_Zero
    if text == "📂 Trimit PUB_Zero":
        USER_STATE[chat_id]["await"] = "PUB_ZERO"
        send_message(chat_id, "Trimite acum fisierul pentru PUB_Zero (Excel).")
        return jsonify(ok=True)

    # Proceseaza
    if text == "🚀 Proceseaza":
        return handle_process(chat_id)

    # Alt text
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
    state = USER_STATE.get(chat_id) or {"await": None}
    role = state.get("await")

    if role not in ("IN", "PUB_ZERO"):
        send_message(
            chat_id,
            "Mai intai apasa „📂 Trimit IN” sau „📂 Trimit PUB_Zero”, apoi trimite fisierul corespunzator.",
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
        dest_path = os.path.join(user_dir, f"IN{ext}")
    else:
        dest_path = os.path.join(user_dir, f"PUB_Zero{ext}")

    ok = download_file(file_id, dest_path)
    if not ok:
        send_message(chat_id, "Nu am reusit sa descarc fisierul. Incearca din nou.")
        return jsonify(ok=True)

    if role == "IN":
        send_message(chat_id, "Am salvat fisierul IN ✅")
    else:
        send_message(chat_id, "Am salvat fisierul PUB_Zero ✅")

    # resetam asteptarea
    USER_STATE[chat_id]["await"] = None

    # daca ambele exista pe disc, il anuntam
    in_exists = find_latest_with_prefix(user_dir, "IN") is not None
    pub_exists = find_latest_with_prefix(user_dir, "PUB_Zero") is not None
    if in_exists and pub_exists:
        send_message(chat_id, "Ambele fisiere sunt pregatite ✅\nApasa „🚀 Proceseaza”.", keyboard=True)

    return jsonify(ok=True)


# ================== HANDLE PROCESS ==================

def handle_process(chat_id: int):
    user_dir = get_user_dir(chat_id)

    # Cautam cele mai noi fisiere IN si PUB_Zero direct pe disc
    in_path = find_latest_with_prefix(user_dir, "IN")
    pub_zero_path = find_latest_with_prefix(user_dir, "PUB_Zero")

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
        # 1) Din PUB_Zero -> PUB_IN (_modificat)
        pub_in_path = format_pub_zero(pub_zero_path)

        # 2) Din IN + PUB_IN -> IN_modificat
        final_path = run_combined_flow(in_path, pub_in_path)

        # Trimitem fisierele inapoi
        send_file(chat_id, pub_in_path, "PUB_IN (din PUB_Zero)")
        send_file(chat_id, final_path, "IN_modificat (din IN + PUB_IN)")

        send_message(chat_id, "Gata ✅ Ti-am trimis fisierele modificate.")

        # NU stergem fisierele imediat – daca vrei, putem adauga buton de reset mai tarziu

    except Exception as e:
        send_message(chat_id, f"A aparut o eroare la procesare:\n{e}", keyboard=True)

    return jsonify(ok=True)


# ================== MAIN ==================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
