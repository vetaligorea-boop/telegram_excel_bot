import os
import requests
from flask import Flask, request, jsonify
from processor import format_pub_zero, run_combined_flow

app = Flask(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Lipseste BOT_TOKEN (seteaza variabila in Render).")

TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

# stare in memorie: pentru fiecare chat_id tinem IN si PUB_Zero
USER_STATE = {}  # chat_id -> {"await": None/"IN"/"PUB_ZERO", "in_path": str, "pub_zero_path": str}


def get_user_dir(chat_id: int) -> str:
    base = "/tmp/telegram_excel_bot"
    path = os.path.join(base, str(chat_id))
    os.makedirs(path, exist_ok=True)
    return path


@app.route("/", methods=["GET"])
def index():
    return "Bot online ✅", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}

    if "message" not in update:
        return jsonify(ok=True)

    message = update["message"]
    chat_id = message["chat"]["id"]

    # init state
    if chat_id not in USER_STATE:
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}

    # document upload
    if "document" in message:
        return handle_document(message, chat_id)

    # text / comenzi
    text = message.get("text", "") or ""

    if text.startswith("/start"):
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}
        send_message(
            chat_id,
            "Salut 👋\n\n"
            "1️⃣ Apasa „📂 Trimit IN” si trimite fisierul IN (.xlsx/.xlsm)\n"
            "2️⃣ Apasa „📂 Trimit PUB_Zero” si trimite fisierul PUB_Zero\n"
            "3️⃣ Apasa „🚀 Proceseaza” ca sa primesti PUB_IN si FINAL."
        , keyboard=True)
        return jsonify(ok=True)

    if text == "📂 Trimit IN":
        USER_STATE[chat_id]["await"] = "IN"
        send_message(chat_id, "Trimite acum fisierul pentru IN (playlist).")
        return jsonify(ok=True)

    if text == "📂 Trimit PUB_Zero":
        USER_STATE[chat_id]["await"] = "PUB_ZERO"
        send_message(chat_id, "Trimite acum fisierul pentru PUB_Zero.")
        return jsonify(ok=True)

    if text == "🚀 Proceseaza":
        return handle_process(chat_id)

    # alt text
    send_message(
        chat_id,
        "Te rog foloseste butoanele:\n"
        "📂 Trimit IN -> apoi trimite fisierul IN\n"
        "📂 Trimit PUB_Zero -> apoi trimite fisierul PUB_Zero\n"
        "🚀 Proceseaza -> pentru rezultat.",
        keyboard=True
    )
    return jsonify(ok=True)


def handle_document(message, chat_id):
    state = USER_STATE.setdefault(chat_id, {"await": None, "in_path": None, "pub_zero_path": None})
    doc = message["document"]
    file_id = doc["file_id"]
    file_name = doc.get("file_name", "fisier.xlsx")

    role = state.get("await")
    if role not in ("IN", "PUB_ZERO"):
        send_message(
            chat_id,
            "Spune-mi intai ce fișier este:\n"
            "Apasa „📂 Trimit IN” sau „📂 Trimit PUB_Zero”, apoi retrimite fisierul.",
            keyboard=True
        )
        return jsonify(ok=True)

    # luam file_path
    r = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id})
    data = r.json()
    if not data.get("ok"):
        send_message(chat_id, "Nu pot descarca fisierul (getFile a esuat).")
        return jsonify(ok=True)

    file_path = data["result"]["file_path"]
    download_url = f"{TELEGRAM_FILE_API}/{file_path}"
    resp = requests.get(download_url)
    if resp.status_code != 200:
        send_message(chat_id, "Nu pot descarca fisierul (download esuat).")
        return jsonify(ok=True)

    # salvam local in folderul utilizatorului
    user_dir = get_user_dir(chat_id)
    ext = os.path.splitext(file_name)[1] or ".xlsx"
    save_name = f"{role}{ext}"
    local_path = os.path.join(user_dir, save_name)
    with open(local_path, "wb") as f:
        f.write(resp.content)

    if role == "IN":
        state["in_path"] = local_path
        send_message(chat_id, "Am salvat fisierul IN ✅")
    elif role == "PUB_ZERO":
        state["pub_zero_path"] = local_path
        send_message(chat_id, "Am salvat fisierul PUB_Zero ✅")

    # dupa upload, resetam "await"
    state["await"] = None

    # daca avem ambele, sugeram procesarea
    if state["in_path"] and state["pub_zero_path"]:
        send_message(chat_id, "Ambele fisiere sunt pregatite ✅\nApasa „🚀 Proceseaza”.", keyboard=True)

    return jsonify(ok=True)


def handle_process(chat_id):
    state = USER_STATE.get(chat_id) or {}
    in_path = state.get("in_path")
    pub_zero_path = state.get("pub_zero_path")

    if not in_path or not pub_zero_path:
        msg = "Inca lipseste ceva:\n"
        if not in_path:
            msg += "- fisierul IN\n"
        if not pub_zero_path:
            msg += "- fisierul PUB_Zero\n"
        msg += "\nFoloseste butoanele pentru a le trimite."
        send_message(chat_id, msg, keyboard=True)
        return jsonify(ok=True)

    send_message(chat_id, "Procesez fisierele... 📊 Te rog asteapta rezultatul.")

    try:
        # 1) PUB_Zero -> PUB_IN
        pub_in_path = format_pub_zero(pub_zero_path)

        # 2) IN + PUB_IN -> FINAL
        final_path = run_combined_flow(in_path, pub_in_path)

        # trimitem PUB_IN
        with open(pub_in_path, "rb") as f:
            files = {"document": (os.path.basename(pub_in_path), f)}
            data = {"chat_id": chat_id}
            requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files)

        # trimitem FINAL
        with open(final_path, "rb") as f:
            files = {"document": (os.path.basename(final_path), f)}
            data = {"chat_id": chat_id}
            requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files)

        send_message(chat_id, "Gata ✅ Ti-am trimis PUB_IN si FINAL.")

        # optional: resetam starea
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}

    except Exception as e:
        send_message(chat_id, f"A aparut o eroare la procesare: {e}", keyboard=True)

    return jsonify(ok=True)


def send_message(chat_id, text, keyboard=False):
    payload = {
        "chat_id": chat_id,
        "text": text
    }
    if keyboard:
        payload["reply_markup"] = {
            "keyboard": [
                [{"text": "📂 Trimit IN"}, {"text": "📂 Trimit PUB_Zero"}],
                [{"text": "🚀 Proceseaza"}],
            ],
            "resize_keyboard": True,
            "one_time_keyboard": False
        }
    requests.post(f"{TELEGRAM_API}/sendMessage", json=payload)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
