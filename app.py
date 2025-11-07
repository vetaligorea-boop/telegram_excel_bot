import os
import requests
from flask import Flask, request, jsonify
from processor import format_pub_zero, run_combined_flow

app = Flask(__name__)

# ================== CONFIG ==================

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Lipseste BOT_TOKEN. Seteaza variabila de mediu in Render.")

TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"

# Pentru fiecare chat pastram doar fisierele lui
# USER_STATE NU atinge nimic in afara fișierelor incarcate de acel utilizator
USER_STATE = {}  # chat_id -> {"await": None / "IN" / "PUB_ZERO", "in_path": str, "pub_zero_path": str}


# ================== HELPERI ==================

def get_user_dir(chat_id: int) -> str:
    """
    Creeaza un folder separat pentru fiecare utilizator:
    /tmp/telegram_excel_bot/<chat_id>/
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
        # nu blocam serverul pe eroare de retea
        pass


def download_file(file_id: str, dest_path: str) -> bool:
    """
    Descarca un fisier trimis pe Telegram in dest_path.
    """
    try:
        r = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id}, timeout=10)
        data = r.json()
        if not data.get("ok"):
            return False

        file_path = data["result"]["file_path"]
        url = f"{TELEGRAM_FILE_API}/{file_path}"

        resp = requests.get(url, timeout=30)
        if resp.status_code != 200:
            return False

        with open(dest_path, "wb") as f:
            f.write(resp.content)

        return True
    except Exception:
        return False


# ================== ROUTE SIMPLU PENTRU TEST ==================

@app.route("/", methods=["GET"])
def index():
    return "Bot online ✅", 200


# ================== WEBHOOK (TELEGRAM) ==================

@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}

    message = update.get("message")
    if not message:
        # ignoram alte tipuri de update (callback, edited_message etc)
        return jsonify(ok=True)

    chat = message.get("chat") or {}
    chat_id = chat.get("id")
    if not chat_id:
        return jsonify(ok=True)

    # initializam stare pentru acest chat daca nu exista
    state = USER_STATE.setdefault(chat_id, {"await": None, "in_path": None, "pub_zero_path": None})

    # 1) daca vine document (fisier Excel)
    if "document" in message:
        return handle_document(message, chat_id, state)

    # 2) daca vine text
    text = (message.get("text") or "").strip()

    # /start
    if text.startswith("/start"):
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}
        send_message(
            chat_id,
            "Salut 👋\n\n"
            "Trimite fisierele DOAR prin aceste butoane:\n"
            "1️⃣ Apasa „📂 Trimit IN” si trimite fisierul IN (.xlsx/.xlsm/.xls)\n"
            "2️⃣ Apasa „📂 Trimit PUB_Zero” si trimite fisierul PUB_Zero\n"
            "3️⃣ Apasa „🚀 Proceseaza” pentru a primi fisierele _modificat.\n\n"
            "Fisierele originale nu sunt modificate, doar copiile trimise inapoi.",
            keyboard=True,
        )
        return jsonify(ok=True)

    # Alege fisierul IN
    if text == "📂 Trimit IN":
        state["await"] = "IN"
        send_message(chat_id, "Trimite acum fisierul pentru IN (playlist).")
        return jsonify(ok=True)

    # Alege fisierul PUB_Zero
    if text == "📂 Trimit PUB_Zero":
        state["await"] = "PUB_ZERO"
        send_message(chat_id, "Trimite acum fisierul pentru PUB_Zero.")
        return jsonify(ok=True)

    # Proceseaza
    if text == "🚀 Proceseaza":
        return handle_process(chat_id, state)

    # Orice alt text
    send_message(
        chat_id,
        "Te rog foloseste butoanele de mai jos:\n"
        "📂 Trimit IN → apoi trimite fisierul IN\n"
        "📂 Trimit PUB_Zero → apoi trimite fisierul PUB_Zero\n"
        "🚀 Proceseaza → pentru rezultat.",
        keyboard=True,
    )
    return jsonify(ok=True)


# ================== HANDLERS ==================

def handle_document(message, chat_id, state):
    """
    Salveaza fisierul trimis ca IN sau PUB_Zero
    DOAR pentru acest chat_id (nu atinge nimic altceva).
    """
    role = state.get("await")
    if role not in ("IN", "PUB_ZERO"):
        send_message(
            chat_id,
            "Mai intai apasa „📂 Trimit IN” sau „📂 Trimit PUB_Zero”, apoi trimite fisierul corect.",
            keyboard=True,
        )
        return jsonify(ok=True)

    doc = message["document"]
    file_id = doc["file_id"]
    file_name = doc.get("file_name", "fisier.xlsx")

    user_dir = get_user_dir(chat_id)
    _, ext = os.path.splitext(file_name)
    if not ext:
        ext = ".xlsx"

    # denumim clar local: IN.ext sau PUB_Zero.ext
    if role == "IN":
        local_path = os.path.join(user_dir, f"IN{ext}")
    else:
        local_path = os.path.join(user_dir, f"PUB_Zero{ext}")

    ok = download_file(file_id, local_path)
    if not ok:
        send_message(chat_id, "Nu pot descarca fisierul de la Telegram. Incearca din nou.")
        state["await"] = None
        return jsonify(ok=True)

    if role == "IN":
        state["in_path"] = local_path
        send_message(chat_id, "Am salvat fisierul IN ✅")
    else:
        state["pub_zero_path"] = local_path
        send_message(chat_id, "Am salvat fisierul PUB_Zero ✅")

    state["await"] = None

    # daca ambele sunt deja incarcate
    if state.get("in_path") and state.get("pub_zero_path"):
        send_message(chat_id, "Ambele fisiere sunt pregatite ✅\nApasa „🚀 Proceseaza”.", keyboard=True)

    return jsonify(ok=True)


def handle_process(chat_id, state):
    """
    Ruleaza format_pub_zero + run_combined_flow DOAR pe fisierele
    incarcate de acest chat. Trimite inapoi copiile _modificat.
    """
    in_path = state.get("in_path")
    pub_zero_path = state.get("pub_zero_path")

    if not in_path or not pub_zero_path:
        msg = "Inca lipsesc fisiere:\n"
        if not in_path:
            msg += "❌ fisierul IN\n"
        if not pub_zero_path:
            msg += "❌ fisierul PUB_Zero\n"
        msg += "\nFoloseste butoanele pentru a le trimite."
        send_message(chat_id, msg, keyboard=True)
        return jsonify(ok=True)

    send_message(chat_id, "Procesez fisierele... ⏳")

    try:
        # 1) Din PUB_Zero -> PUB_IN (PUB_Zero_modificat)
        pub_in_path = format_pub_zero(pub_zero_path)

        # 2) Din IN + PUB_IN -> IN_modificat
        final_path = run_combined_flow(in_path, pub_in_path)

        # Trimitem PUB_IN
        try:
            with open(pub_in_path, "rb") as f:
                files = {"document": (os.path.basename(pub_in_path), f)}
                data = {"chat_id": chat_id}
                requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files, timeout=30)
        except Exception:
            send_message(chat_id, "Nu am reusit sa trimit fisierul PUB_IN_modificat.")

        # Trimitem IN_modificat (FINAL)
        try:
            with open(final_path, "rb") as f:
                files = {"document": (os.path.basename(final_path), f)}
                data = {"chat_id": chat_id}
                requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files, timeout=30)
        except Exception:
            send_message(chat_id, "Nu am reusit sa trimit fisierul FINAL_modificat.")

        send_message(chat_id, "Gata ✅ Ti-am trimis fisierele _modificat.")

        # Resetam starea ca sa poti incarca alt set de fisiere
        USER_STATE[chat_id] = {"await": None, "in_path": None, "pub_zero_path": None}

    except Exception as e:
        # In caz de eroare, nu crapam, doar anuntam
        send_message(chat_id, f"A aparut o eroare la procesare: {e}", keyboard=True)

    return jsonify(ok=True)


# ================== MAIN (pentru rulare locala) ==================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
