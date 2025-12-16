import asyncio
import os
import shutil
import tempfile
from pathlib import Path

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiohttp import web

# ====== CONFIG ======
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN lipseste din Render -> Environment Variables")

PORT = int(os.getenv("PORT", "10000"))

# ====== TELEGRAM BOT ======
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Render Free: procesam doar .xlsx si .xlsm (fara conversie .xls)
ALLOWED_EXT = {".xlsx", ".xlsm"}


def process_excel(path: Path) -> Path:
    """
    Aici va veni logica ta reala (macro-ul VBA transpus in Python).
    Acum: doar copiaza fisierul ca test ca botul merge cap-coada.
    """
    out = path.with_name(path.stem + "_modificat" + path.suffix)
    shutil.copy(path, out)
    return out


@dp.message(CommandStart())
async def start(msg: types.Message):
    await msg.answer(
        "✅ Sunt online.\n"
        "Trimite-mi un fișier Excel (.xlsx sau .xlsm) și ți-l procesez.\n"
        "⚠️ Formatul .xls nu este suportat aici (te rog salvează-l ca .xlsx)."
    )


@dp.message(F.document)
async def handle_file(msg: types.Message):
    doc = msg.document
    filename = doc.file_name or "fisier.xlsx"
    ext = Path(filename).suffix.lower()

    # Daca e .xls, explicam ce sa faca
    if ext == ".xls":
        await msg.answer(
            "❌ Fișierul tău este .xls (format vechi).\n"
            "Pe Render Free nu pot converti .xls → .xlsx.\n"
            "✅ Te rog: deschide fișierul în Excel → Save As → .xlsx și retrimite."
        )
        return

    if ext not in ALLOWED_EXT:
        await msg.answer("❌ Format neacceptat. Trimite doar .xlsx sau .xlsm.")
        return

    await msg.answer("📥 Am primit fișierul. Îl descarc...")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        in_file = tmpdir / filename

        # descarcare din Telegram
        file = await bot.get_file(doc.file_id)
        await bot.download_file(file.file_path, in_file)

        await msg.answer("⚙️ Procesez fișierul ...")
        try:
            out_file = process_excel(in_file)
        except Exception as e:
            await msg.answer(f"❌ Eroare la procesare: {e}")
            return

        await msg.answer_document(types.FSInputFile(out_file), caption="✅ Gata. Iată fișierul procesat.")


# ====== HEALTH SERVER (pentru Render Web Service) ======
async def run_health_server():
    app = web.Application()

    async def health(_):
        return web.Response(text="OK")

    app.router.add_get("/", health)
    app.router.add_get("/health", health)

    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()


async def main():
    # deschidem portul ca Render sa nu opreasca serviciul
    await run_health_server()
    # pornim botul
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
