import os
import tempfile
import requests
from flask import Flask, request, jsonify
from processor import process_workbook

app = Flask(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("Lipseste BOT_TOKEN (seteaza variabila de mediu in hosting).")

TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TELEGRAM_FILE_API = f"https://api.telegram.org/file/bot{BOT_TOKEN}"


@app.route("/", methods=["GET"])
def index():
    return "Bot online ✅", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}

    if "message" in update:
        message = update["message"]
        chat_id = message["chat"]["id"]

        if "document" in message:
            return handle_document(message, chat_id)

        text = message.get("text", "")
        if text.startswith("/start"):
            send_message(chat_id,
                         "Salut! Trimite-mi un fișier .xlsx sau .xlsm și îți trimit înapoi versiunea _modificat.")
        else:
            send_message(chat_id,
                         "Trimite-mi un fișier Excel (.xlsx / .xlsm) și îl procesez pentru tine.")
        return jsonify(ok=True)

    return jsonify(ok=True)


def handle_document(message, chat_id):
    doc = message["document"]
    file_id = doc["file_id"]
    file_name = doc.get("file_name", "fisier.xlsx")

    if not (file_name.lower().endswith(".xlsx") or file_name.lower().endswith(".xlsm")):
        send_message(chat_id, "Te rog trimite un fișier .xlsx sau .xlsm.")
        return jsonify(ok=True)

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

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file_name)
        with open(input_path, "wb") as f:
            f.write(resp.content)

        try:
            output_path = process_workbook(input_path)
        except Exception as e:
            print("Eroare la procesare:", e)
            send_message(chat_id, f"Eroare la procesare: {e}")
            return jsonify(ok=True)

        with open(output_path, "rb") as f:
            files = {"document": (os.path.basename(output_path), f)}
            data = {"chat_id": chat_id}
            r = requests.post(f"{TELEGRAM_API}/sendDocument", data=data, files=files)

        if not r.ok:
            send_message(chat_id, "A apărut o eroare la trimiterea fișierului modificat.")

    return jsonify(ok=True)


def send_message(chat_id, text):
    requests.post(f"{TELEGRAM_API}/sendMessage", json={
        "chat_id": chat_id,
        "text": text
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
