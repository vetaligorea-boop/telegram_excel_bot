import os
import requests
from flask import Flask, request, jsonify

from processor import format_pub_zero, run_combined_flow

# ================== CONFIG ==================

# !!! INLOCUIESTE cu tokenul tau real de la BotFather !!!
BOT_TOKEN = "PUUNE_TOKENUL_TAU_AICI"
API_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"
FILE_API_URL = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

ALLOWED_EXTS = [".xls", ".xlsx", ".xlsm"]

app = Flask(__name__)

# memorie simpla in RAM pe baza de chat_id
USER_STATE = {}


# ================== HELPERI ==================

def get_user_state(chat_id):
    """
    Returneaza / initializeaza dictionarul de stare pentru fiecare utilizator.
    Folosim string ca si cheie ca sa fie mereu consistent.
    """
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
    payload = {
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "HTML",
    }
    if reply_markup:
        payload["reply_markup"] = reply_markup
    requests.post(f"{API_URL}/sendMessage", json=payload)


def send_document(chat_id, file_path, caption=None):
    if not os.path.isfile(file_path):
        send_message(chat_id, f"Nu gasesc fisierul: {file_path}")
        return
    with open(file_path, "rb") as f:
        files = {"document": f}
        data = {
            "chat_id": chat_id,
        }
        if caption:
            data["caption"] = caption
        requests.post(f"{API_URL}/sendDocument", data=data, files=files)


def get_file(file_id):
    resp = requests.get(f"{API_URL}/getFile", params={"file_id": file_id})
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"Nu pot obtine file_path pentru {file_id}: {data}")
    return data["result"]


def main_keyboard():
    """
    Inline keyboard cu cele 3 butoane: IN, PUB_Zero, Proceseaza
    """
    return {
        "inline_keyboard": [
            [
                {"text": "📂 Trimit IN", "callback_data": "SEND_IN"},
                {"text": "📂 Trimit PUB_Zero", "callback_data": "SEND_PUB_ZERO"},
            ],
            [
                {"text": "🚀 Proceseaza", "callback_data": "PROCESS"},
            ],
        ]
    }


# ================== HANDLER: DOCUMENT ==================

def handle_document(message):
    chat_id = message["chat"]["id"]
    state = get_user_state(chat_id)

    role = state.get("await")
    if role not in ("IN", "PUB_ZERO"):
        send_message(chat_id, "Te rog apasa mai intai butonul pentru IN sau PUB_Zero.", main_keyboard())
        return

    if "document" not in message:
        send_message(chat_id, "Nu vad niciun fisier atasat.")
        return

    file_id = message["document"]["file_id"]
    file_name = message["document"]["file_name"]

    _, ext = os.path.splitext(file_name)
    ext = ext.lower()

    if ext not in ALLOWED_EXTS:
        send_message(chat_id, "Accept doar fisiere .xls, .xlsx sau .xlsm.")
        return

    # luam file_path de la Telegram
    info = get_file(file_id)
    file_url = f"{FILE_API_URL}/{info['file_path']}"

    user_dir = ensure_user_dir(chat_id)

    # pastram numele ORIGINAL
    save_name = file_name
    local_path = os.path.join(user_dir, save_name)

    resp = requests.get(file_url)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(resp.content)

    if role == "IN":
        state["in_path"] = local_path
        send_message(chat_id, "Am salvat fisierul IN ✅", main_keyboard())
    elif role == "PUB_ZERO":
        state["pub_zero_path"] = local_path
        send_message(chat_id, "Am salvat fisierul PUB_Zero ✅", main_keyboard())

    # resetam asteptarea
    state["await"] = None

    # daca avem ambele, informam utilizatorul
    if state.get("in_path") and state.get("pub_zero_path"):
        send_message(
            chat_id,
            "Ambele fisiere sunt pregatite ✅\nApasa „🚀 Proceseaza”.",
            main_keyboard()
        )


# ================== HANDLER: CALLBACK (BUTOANE) ==================

def handle_callback_query(callback):
    data = callback.get("data")
    message = callback.get("message", {})
    chat_id = message["chat"]["id"]

    state = get_user_state(chat_id)

    if data == "SEND_IN":
        state["await"] = "IN"
        send_message(chat_id, "Trimite acum fisierul pentru IN (playlist).")

    elif data == "SEND_PUB_ZERO":
        state["await"] = "PUB_ZERO"
        send_message(chat_id, "Trimite acum fisierul pentru PUB_Zero.")

    elif data == "PROCESS":
        in_path = state.get("in_path")
        pub_zero_path = state.get("pub_zero_path")

        missing = []
        if not in_path:
            missing.append("- fisierul IN")
        if not pub_zero_path:
            missing.append("- fisierul PUB_Zero")

        if missing:
            send_message(
                chat_id,
                "Inca lipsesc fisiere:\n" + "\n".join(missing) + "\n\nFoloseste butoanele pentru a le trimite.",
                main_keyboard()
            )
            return

        send_message(chat_id, "Procesez fisierele... 📊")

        try:
            # 1) Generam PUB_IN din PUB_Zero (doar valorile necesare, formatul ramane)
            pub_in_path = format_pub_zero(pub_zero_path)

            # 2) Generam FINAL din IN + PUB_IN
            final_path = run_combined_flow(in_path, pub_in_path)

            # trimitem PUB_IN
            send_document(chat_id, pub_in_path, "PUB_IN (generat din PUB_Zero)")

            # trimitem FINAL
            send_document(chat_id, final_path, "Desfasurator FINAL ✅")

        except Exception as e:
            send_message(chat_id, f"Eroare la procesare: {e}", main_keyboard())


# ================== WEBHOOK ROUTES ==================

@app.route("/", methods=["GET"])
def index():
    return "Bot online ✅"


@app.route("/", methods=["POST"])
def webhook():
    update = request.get_json(force=True)

    # mesaje normale
    if "message" in update:
        msg = update["message"]
        chat_id = msg["chat"]["id"]
        text = msg.get("text", "")

        # START
        if text == "/start":
            # resetam starea cand incepe
            USER_STATE[str(chat_id)] = {}
            send_message(
                chat_id,
                "Salut 👋\n\n"
                "1️⃣ Apasa „📂 Trimit IN” si trimite fisierul IN (.xls/.xlsx/.xlsm)\n"
                "2️⃣ Apasa „📂 Trimit PUB_Zero” si trimite fisierul PUB_Zero\n"
                "3️⃣ Apasa „🚀 Proceseaza” ca sa primesti PUB_IN si FINAL.",
                main_keyboard()
            )
        # DOCUMENT
        elif "document" in msg:
            handle_document(msg)
        # orice alt text
        else:
            send_message(
                chat_id,
                "Foloseste butoanele de mai jos pentru a trimite fisierele si a procesa.",
                main_keyboard()
            )

    # callback de la butoane
    if "callback_query" in update:
        handle_callback_query(update["callback_query"])

    return jsonify(ok=True)
