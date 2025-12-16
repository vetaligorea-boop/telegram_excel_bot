import asyncio
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart

BOT_TOKEN = os.getenv("BOT_TOKEN")

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
    return input_path.with_suffix(".xlsx")


def process_excel(path: Path) -> Path:
    out = path.with_name(path.stem + "_modificat" + path.suffix)
    shutil.copy(path, out)
    return out


@dp.message(CommandStart())
async def start(msg: types.Message):
    await msg.answer("Trimite un fișier Excel (.xls, .xlsx, .xlsm)")


@dp.message(F.document)
async def handle_file(msg: types.Message):
    doc = msg.document
    ext = Path(doc.file_name).suffix.lower()

    if ext not in ALLOWED_EXT:
        await msg.answer("Format Excel neacceptat.")
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        in_file = tmp / doc.file_name

        file = await bot.get_file(doc.file_id)
        await bot.download_file(file.file_path, in_file)

        work = in_file
        if ext == ".xls":
            await msg.answer("Convertesc .xls → .xlsx")
            work = convert_xls_to_xlsx(in_file)

        await msg.answer("Procesez fișierul…")
        out = process_excel(work)

        await msg.answer_document(types.FSInputFile(out))


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
