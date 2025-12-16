import asyncio
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiohttp import web

BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN lipseste din Render -> Environment Variables")

PORT = int(os.getenv("PORT", "10000"))

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

ALLOWED_EXT = {".xls", ".xlsx", ".xlsm"}


def convert_xls_to_xlsx(input_path: Path) -> Path:
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "xlsx",
        str(input_path),
        "--outdir", str(input_path.parent)
    ]
    subprocess.run(cmd, check=True)
    return input_path.parent / (input_path.stem + ".xlsx")


def process_excel(path: Path) -> Path:
    out = path.with_name(path.stem + "_modificat" + path.suffix)
    shutil.copy(path, out)
    return out


@dp.message(CommandStart())
async def start(msg: types.Message):
    await msg.answer("✅ Sunt online. Trimite un fișier Excel (.xls, .xlsx, .xlsm).")


@dp.message(F.document)
async def handle_file(msg: types.Message):
    doc = msg.document
    ext = Path(doc.file_name).suffix.lower()

    if ext not in ALLOWED_EXT:
        await msg.answer("❌ Format neacceptat. Trimite .xls / .xlsx / .xlsm")
        return

    await msg.answer("📥 Am primit fișierul. Îl descarc...")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        in_file = tmp / doc.file_name

        file = await bot.get_file(doc.file_id)
        await bot.download_file(file.file_path, in_file)

        work = in_file
        if ext == ".xls":
            await msg.answer("🔁 Convertesc .xls → .xlsx ...")
            work = convert_xls_to_xlsx(in_file)

        await msg.answer("⚙️ Procesez fișierul ...")
        out = process_excel(work)

        await msg.answer_document(types.FSInputFile(out), caption="✅ Gata.")


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
    # important: deschidem portul pentru Render
    await run_health_server()
    # apoi pornim botul
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
