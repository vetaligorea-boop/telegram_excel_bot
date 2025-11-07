import os
import requests
from flask import Flask, request, jsonify
from processor import process_pair

app = Flask(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Lipseste BOT_TOKEN (seteaza in Render).")

TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

# memorie simpla per user: ce fisier a trimis
SESSIONS = {}
BASE_TMP = "/tmp/telegram_excel_bot"
os.makedirs(BASE_TMP, exist_ok=True)


def get_session(user_id):
    if user_id not in SESSIONS:
        SESSIONS[user_id] = {
            "in_path": None,
            "pub_path": None,
            "waiting": None,  # "in" sau "pub"
        }
    return SESSIONS[user_id]


def build_menu():
    return {
        "inline_keyboard": [
            [
                {"text": "1️⃣ Trimite fișier IN", "callback_data": "upload_in"},
                {"text": "2️⃣ Trimite fișier PUB_Zero", "callback_data": "upload_pub"},
            ],
            [
                {"text": "▶️ Procesează", "callback_data": "process"},
                {"text": "♻️ Reset", "callback_data": "reset"},
            ],
        ]
    }


def send_message(chat_id, text, reply_markup=None):
    payload = {"chat_id": chat_id, "text": text}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    requests.post(f"{TELEGRAM_API}/sendMessage", json=payload)


@app.route("/", methods=["GET"])
def index():
    return "Bot online ✅", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}

    # butoane (inline keyboard)
    if "callback_query" in update:
        return handle_callback(update["callback_query"])

    # mesaje normale (text / document)
    if "message" in update:
        message = update["message"]
        chat_id = message["chat"]["id"]
        user_id = message["from"]["id"]

        # daca e fisier
        if "document" in message:
            return handle_document(message, chat_id, user_id)

        # daca e text
        text = message.get("text", "")
        if text.startswith("/start"):
            SESSIONS[user_id] = {"in_path": None, "pub_path": None, "waiting": None}
            send_message(
                chat_id,
                "Salut! 👋\n\n"
                "1️⃣ Apasă „Trimite fișier IN”, apoi trimite fișierul Excel pentru IN.\n"
                "2️⃣ Apasă „Trimite fișier PUB_Zero”, apoi trimite fișierul Excel pentru PUB_Zero.\n"
                "3️⃣ Apasă „Procesează” ca să primești fișierul final (_modificat).",
                reply_markup=build_menu(),
            )
        else:
            send_message(
                chat_id,
                "Folosește butoanele de mai jos:\n"
                "1️⃣ Trimite fișier IN\n2️⃣ Trimite fișier PUB_Zero\n▶️ Procesează",
                reply_markup=build_menu(),
            )

    return jsonify(ok=True)


def handle_callback(cb):
    chat_id = cb["message"]["chat"]["id"]
    user_id = cb["from"]["id"]
    data = cb.get("data")
    session = get_session(user_id)

    # raspunde ca sa dispara animatia de loading
    requests.post(f"{TELEGRAM_API}/answerCallbackQuery",
                  json={"callback_query_id": cb["id"]})

    if data == "upload_in":
        session["waiting"] = "in"
        send_message(chat_id,
                     "📂 Acum trimite fișierul Excel corespunzător mapei IN (.xlsx / .xlsm).")
    elif data == "upload_pub":
        session["waiting"] = "pub"
        send_message(chat_id,
                     "📂 Acum trimite fișierul Excel corespunzător mapei PUB_Zero (.xlsx / .xlsm).")
    elif data == "reset":
        SESSIONS[user_id] = {"in_path": None, "pub_path": None, "waiting": None}
        send_message(chat_id,
                     "Session resetată. Poți începe din nou 👇",
                     reply_markup=build_menu())
    elif data == "process":
        in_path = session.get("in_path")
        pub_path = session.get("pub_path")
        if not in_path and not pub_path:
            send_message(chat_id,
                         "Nu am niciun fișier. Începe cu butoanele și trimite fișierele.",
                         reply_markup=build_menu())
        elif not in_path:
            send_message(chat_id,
                         "Îmi lipsește fișierul IN. Apasă „Trimite fișier IN” și trimite-l.",
                         reply_markup=build_menu())
        elif not pub_path:
            send_message(chat_id,
                         "Îmi lipsește fișierul PUB_Zero. Apasă „Trimite fișier PUB_Zero” și trimite-l.",
                         reply_markup=build_menu())
        else:
            try:
                send_message(chat_id, "Procesăm fișierele... 🔄 Așteaptă rezultatul.")
                final_path = process_pair(in_path, pub_path)
            except Exception as e:
                print("Eroare la process_pair:", e)
                send_message(chat_id, f"Eroare la procesare: {e}")
                return jsonify(ok=True)

            # trimitem fisierul final
            try:
                with open(final_path, "rb") as f:
                    files = {"document": (os.path.basename(final_path), f)}
                    data = {"chat_id": chat_id}
                    r = requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files)
                if r.ok:
                    send_message(chat_id,
                                 "✅ Gata! Acesta este fișierul FINAL (_modificat).",
                                 reply_markup=build_menu())
                else:
                    send_message(chat_id,
                                 "Am generat fișierul, dar nu am reușit să-l trimit 😕")
            except Exception as e:
                print("Eroare la trimitere:", e)
                send_message(chat_id, f"Eroare la trimitere: {e}")

    return jsonify(ok=True)


def handle_document(message, chat_id, user_id):
    session = get_session(user_id)
    target = session.get("waiting")

    if not target:
        # daca nu a ales ce tip e fisierul
        send_message(
            chat_id,
            "Te rog întâi alege cu butonul dacă acest fișier este pentru IN sau pentru PUB_Zero.",
            reply_markup=build_menu(),
        )
        return jsonify(ok=True)

    doc = message["document"]
    file_id = doc["file_id"]
    file_name = doc.get("file_name", "fisier.xlsx").lower()

    if not (file_name.endswith(".xlsx") or file_name.endswith(".xlsm")):
        send_message(chat_id, "Accept doar .xlsx sau .xlsm.")
        return jsonify(ok=True)

    # luam file_path
    r = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id})
    data = r.json()
    if not data.get("ok"):
        send_message(chat_id, "Nu pot descărca fișierul (getFile a eșuat).")
        return jsonify(ok=True)

    file_path = data["result"]["file_path"]
    download_url = f"{TELEGRAM_FILE_API}/{file_path}"

    resp = requests.get(download_url)
    if resp.status_code != 200:
        send_message(chat_id, "Nu pot descărca fișierul (download eșuat).")
        return jsonify(ok=True)

    # salvam in /tmp per user si tip
    ext = ".xlsx"
    if file_name.endswith(".xlsm"):
        ext = ".xlsm"

    local_name = f"user{user_id}_{target}{ext}"
    local_path = os.path.join(BASE_TMP, local_name)
    with open(local_path, "wb") as f:
        f.write(resp.content)

    if target == "in":
        session["in_path"] = local_path
        msg = "✅ Am salvat fișierul IN."
    else:
        session["pub_path"] = local_path
        msg = "✅ Am salvat fișierul PUB_Zero."

    session["waiting"] = None

    # verificam daca avem ambele
    if session.get("in_path") and session.get("pub_path"):
        msg += "\nPoți apăsa acum „▶️ Procesează” ca să generezi fișierul final."
    else:
        msg += "\nAcum selectează și celălalt tip de fișier din butoane."

    send_message(chat_id, msg, reply_markup=build_menu())
    return jsonify(ok=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
