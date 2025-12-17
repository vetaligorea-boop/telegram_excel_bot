import asyncio
import os
import tempfile
from pathlib import Path
from datetime import datetime, time as dtime

from aiogram import Bot, Dispatcher, F, types
from aiogram.filters import CommandStart
from aiohttp import web

import openpyxl
from openpyxl.styles import Font

# =========================
# CONFIG
# =========================
BOT_TOKEN = os.getenv("BOT_TOKEN")
if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN lipseste din Render -> Environment Variables")

PORT = int(os.getenv("PORT", "10000"))

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

ALLOWED_EXT = {".xlsx", ".xlsm"}  # .xls nu e suportat pe Render Free
WHITE_FONT = Font(color="FFFFFFFF")  # alb (ARGB)

# =========================
# HELPERS
# =========================
def safe_str(v) -> str:
    return "" if v is None else str(v)

def like_prefix(s: str, prefix: str) -> bool:
    return safe_str(s).startswith(prefix)

def set_cell(ws, row: int, col: int, value):
    c = ws.cell(row, col)
    c.value = value
    c.font = WHITE_FONT

def get_last_row_in_col(ws, col: int) -> int:
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(r, col).value
        if v not in (None, ""):
            return r
    return 1

def get_fill_rgb_safe(cell):
    try:
        f = cell.fill
        if not f or not f.patternType:
            return None
        return f.fgColor.rgb
    except Exception:
        return None

def is_yellow_colorindex6(cell) -> bool:
    rgb = get_fill_rgb_safe(cell)
    return rgb in ("FFFFFF00", "FFFF00")

def format_time_value(v) -> str:
    """
    Converteste valoarea din Excel (coloana ORA) in text HH:MM:SS.
    - datetime/time -> ok
    - numar (fractia din zi) -> convertim
    - text "06:30:00" -> pastram
    """
    if v is None or v == "":
        return ""

    if isinstance(v, datetime):
        return v.strftime("%H:%M:%S")

    if isinstance(v, dtime):
        return v.strftime("%H:%M:%S")

    if isinstance(v, (int, float)):
        frac = float(v) % 1.0
        total_seconds = int(round(frac * 86400)) % 86400
        hh = total_seconds // 3600
        mm = (total_seconds % 3600) // 60
        ss = total_seconds % 60
        return f"{hh:02d}:{mm:02d}:{ss:02d}"

    s = safe_str(v).strip()
    if s == "0":
        return "00:00:00"
    return s

# =========================
# VBA logic (partial, fara insert_rows)
# =========================
def CopiereAutomataCombinata(ws):
    lastRow = get_last_row_in_col(ws, 4)
    insidePlaylist = False

    for i in range(1, lastRow + 1):
        cell_d = ws.cell(i, 4)

        if not is_yellow_colorindex6(cell_d):
            cellValue = safe_str(cell_d.value)
            col6 = safe_str(ws.cell(i, 6).value)
            col3 = ws.cell(i, 3).value

            if (cellValue.startswith("ID PUB_") or cellValue.startswith("ID_PUB_") or cellValue.startswith("ID PUB")):
                if not (like_prefix(col6, "PLAYLIST_IN_") or like_prefix(col6, "PLAYLIST_OUT_")):
                    set_cell(ws, i, 19, cellValue)

            elif col3 not in (None, ""):
                set_cell(ws, i, 21, cellValue)

            else:
                if safe_str(ws.cell(i, 21).value) == "":
                    set_cell(ws, i, 19, cellValue)

        col6_val = safe_str(ws.cell(i, 6).value)
        if like_prefix(col6_val, "PLAYLIST_IN_"):
            insidePlaylist = True
        elif like_prefix(col6_val, "PLAYLIST_OUT_"):
            insidePlaylist = False
        elif insidePlaylist and col6_val != "":
            if safe_str(ws.cell(i, 19).value) == "" and safe_str(ws.cell(i, 21).value) == "":
                set_cell(ws, i, 20, col6_val)
        else:
            ws.cell(i, 20).value = None

        dval = safe_str(ws.cell(i, 4).value)
        if dval in ("CCA_SARE_ZAHAR_GRASIMI", "CCA_ORELE_MESEI", "CCA_GIMNASTICA", "CCA_FRUCTE", "CCA_BEA_APA"):
            set_cell(ws, i, 20, dval)
            ws.cell(i, 19).value = None

        col6_val2 = safe_str(ws.cell(i, 6).value)
        if like_prefix(col6_val2, "PLAYLIST_IN_") or like_prefix(col6_val2, "PLAYLIST_OUT_"):
            set_cell(ws, i, 20, safe_str(ws.cell(i, 4).value))
            ws.cell(i, 19).value = None

def ActualizareColoana23(ws):
    lastRow = get_last_row_in_col(ws, 6)
    for i in range(1, lastRow + 1):
        col6 = safe_str(ws.cell(i, 6).value)

        if like_prefix(col6, "PLAYLIST_IN_"):
            set_cell(ws, i, 23, "pub_start   #COLOR 65535")
        elif like_prefix(col6, "PLAYLIST_OUT_"):
            set_cell(ws, i, 23, "pub_stop   #COLOR 4227327")
        else:
            v = safe_str(ws.cell(i, 23).value).strip().lower()
            if v == "":
                continue
            if v == "ceas":
                set_cell(ws, i, 23, "ceas   #COLOR 8421631")
            elif v == "ap":
                set_cell(ws, i, 23, "ap   #COLOR 8454016")
            elif v in ("cr+ap", "cr+12", "cr+15"):
                set_cell(ws, i, 23, f"{v}   #COLOR 16777088")
            elif v in ("reluare_in", "reluare_mid", "reluare_out"):
                set_cell(ws, i, 23, f"{v}   #COLOR 16744448")
            elif v in ("premiera_in", "premiera_mid", "premiera_out", "premiera_ap", "premiera_12", "premiera_15"):
                set_cell(ws, i, 23, f"{v}   #COLOR 8388863")
            else:
                ws.cell(i, 23).font = WHITE_FONT

def AdaugaCategoryDinColoana21(ws):
    exclude = {"FILLER", "Ceas + Direct", "CEC"}
    lastRow = get_last_row_in_col(ws, 21)
    for i in range(1, lastRow + 1):
        col21 = ws.cell(i, 21).value
        col3 = safe_str(ws.cell(i, 3).value)
        if col21 not in (None, "") and col3 not in exclude:
            set_cell(ws, i, 23, col3)

# =========================
# constant sheet fast build (ORA citita din data_only sheet!)
# =========================
def build_constant_fast(wb, ws, ws_values):
    if "constant" in wb.sheetnames:
        constant = wb["constant"]
        if constant.max_row > 0:
            constant.delete_rows(1, constant.max_row)
    else:
        constant = wb.create_sheet("constant")

    sheets = wb._sheets
    if constant in sheets:
        sheets.remove(constant)
        sheets.append(constant)

    lastRow = max(get_last_row_in_col(ws, 6), get_last_row_in_col(ws, 4))

    exclude_contains = [
        "Jurnalfinanciar", "Jurnalulfinanciar", "PROMO YOUTUBE", "YOUTUBE SOFIA ",
        "JurnalSportivNEW", "JurnalulSportiv NEW", "JurnalSportiv ", "JurnalulSportiv ",
        "Carton ap", "Carton 12", "Carton 15", "INTERZIS_AP", "INTERZIS_12", "INTERZIS_15",
        "EarthTV", "stirile deficienta auz 17 HD", "MeteoNEW", "PostScriptumNEW",
        "Promourile", "bumper", "Studio comentarii", "Fotbal Repriza 1", "Stiri 10 min",
        "Fotbal Repriza 2", "REZUMATE", "DE FACTO Cioban", "DE FACTO Ciobanu",
        "PLANUL EUROPA CIOBANU", "PLANUL EUROPA CIOBAN", "DeFacto Tulbure", "PLANUL EUROPA Tulbure",
        "ID_Promo_", "INTERZIS_", "ID_PROMO_", "ID PROMO_", "ID_PROMO",
        "ID_PUB_", "ID PUB_", "ID_PUB", "____PROMOURIIII___"
    ]

    def excluded_by_list(text: str) -> bool:
        low = (text or "").lower()
        for ex in exclude_contains:
            if ex.lower() in low:
                return True
        return False

    def bad_prefix_19(s: str) -> bool:
        s = s or ""
        return (
            s.startswith("ID PROMO_") or s.startswith("ID_PROMO_") or s.startswith("ID PROMO") or
            s.startswith("ID PUB_") or s.startswith("ID_PUB_") or s.startswith("ID PUB")
        )

    out_row = 1
    insidePlaylist = False
    firstValueFound = False

    def write_row(r19, r20, r21, r23, ora_value):
        nonlocal out_row

        r19s = safe_str(r19).strip()
        if r19s == "____PROMOURIIII___":
            constant.cell(out_row, 119).value = None
        else:
            set_cell(constant, out_row, 119, r19s)

        set_cell(constant, out_row, 120, safe_str(r20))
        set_cell(constant, out_row, 121, safe_str(r21))
        set_cell(constant, out_row, 123, safe_str(r23))

        # ✅ ORA corecta (din ws_values data_only=True)
        set_cell(constant, out_row, 122, format_time_value(ora_value))

        v119 = safe_str(constant.cell(out_row, 119).value)
        v121 = safe_str(constant.cell(out_row, 121).value)
        set_cell(constant, out_row, 125, f"{v119} {v121}".strip())

        out_row += 1

    for i in range(1, lastRow + 1):
        col6 = safe_str(ws.cell(i, 6).value)

        # ✅ ora citita din workbook data_only=True
        ora_value = ws_values.cell(i, 2).value

        if like_prefix(col6, "PLAYLIST_IN_"):
            insidePlaylist = True
            firstValueFound = False
            write_row(ws.cell(i, 19).value, ws.cell(i, 20).value, ws.cell(i, 21).value, ws.cell(i, 23).value, ora_value)
            continue

        if like_prefix(col6, "PLAYLIST_OUT_"):
            insidePlaylist = False
            write_row(ws.cell(i, 19).value, ws.cell(i, 20).value, ws.cell(i, 21).value, ws.cell(i, 23).value, ora_value)
            continue

        if insidePlaylist:
            if not firstValueFound:
                firstValueFound = True
            else:
                write_row("", "ID CUB_PUB_TEST", "", "", ora_value)

        write_row(ws.cell(i, 19).value, ws.cell(i, 20).value, ws.cell(i, 21).value, ws.cell(i, 23).value, ora_value)

        cell19 = safe_str(ws.cell(i, 19).value).strip()
        if cell19 != "" and (not bad_prefix_19(cell19)) and (not excluded_by_list(cell19)):
            write_row("ID CUB_PUB_TEST", "", "", "", ora_value)

    return constant

def AdaugaEventNote(constant):
    lastRow = get_last_row_in_col(constant, 119)

    eventNoteAdded119 = False
    for i in range(1, lastRow + 1):
        v119 = safe_str(constant.cell(i, 119).value)
        v122 = safe_str(constant.cell(i, 122).value)
        if v119 != "" and not eventNoteAdded119:
            set_cell(constant, i, 124, f"{v122} {v119}".strip())
            eventNoteAdded119 = True
        elif v119 == "":
            eventNoteAdded119 = False

    eventNoteAdded120 = False
    for i in range(1, lastRow + 1):
        v120 = safe_str(constant.cell(i, 120).value)
        v122 = safe_str(constant.cell(i, 122).value)
        if v120 != "" and not eventNoteAdded120:
            set_cell(constant, i, 124, f"{v122} {v120}".strip())
            eventNoteAdded120 = True
        elif v120 == "":
            eventNoteAdded120 = False

    for i in range(1, lastRow + 1):
        v121 = safe_str(constant.cell(i, 121).value)
        v122 = safe_str(constant.cell(i, 122).value)
        if v121 != "":
            set_cell(constant, i, 124, f"{v122} {v121}".strip())

def UnireColoane119Si121(constant):
    lastRow = get_last_row_in_col(constant, 119)
    for i in range(1, lastRow + 1):
        v119 = safe_str(constant.cell(i, 119).value)
        v121 = safe_str(constant.cell(i, 121).value)
        set_cell(constant, i, 125, f"{v119} {v121}".strip())

def ActualizareColoana123(constant):
    lastRow = get_last_row_in_col(constant, 123)
    exclude = {"ceas+direct", "ceas + direct", "."}

    for i in range(1, lastRow + 1):
        v = safe_str(constant.cell(i, 123).value).strip()
        low = v.lower()

        if low in exclude:
            constant.cell(i, 123).value = None
            continue

        if low == "pub_start":
            set_cell(constant, i, 123, "pub_start   #COLOR 65535")
        elif low == "pub_stop":
            set_cell(constant, i, 123, "pub_stop   #COLOR 4227327")
        elif low == "ceas":
            set_cell(constant, i, 123, "ceas   #COLOR 8421631")
        elif low == "ap":
            set_cell(constant, i, 123, "ap   #COLOR 8454016")
        elif low in ("cr+ap", "cr+12", "cr+15"):
            set_cell(constant, i, 123, f"{low}   #COLOR 16777088")
        elif low in ("reluare_in", "reluare_mid", "reluare_out"):
            set_cell(constant, i, 123, f"{low}   #COLOR 16744448")
        elif low in ("premiera_in", "premiera_mid", "premiera_out", "premiera_ap", "premiera_12", "premiera_15"):
            set_cell(constant, i, 123, f"{low}   #COLOR 8388863")
        else:
            if constant.cell(i, 123).value not in (None, ""):
                constant.cell(i, 123).font = WHITE_FONT

def clear_sheet1_cols(ws):
    lastRow = get_last_row_in_col(ws, 4)
    for r in range(1, lastRow + 1):
        ws.cell(r, 19).value = None
        ws.cell(r, 20).value = None
        ws.cell(r, 21).value = None
        ws.cell(r, 23).value = None

# =========================
# MAIN PROCESS
# =========================
def process_excel(path: Path) -> Path:
    keep_vba = (path.suffix.lower() == ".xlsm")

    # 1) workbook principal (pastreaza formule + VBA) -> asta salvam
    wb = openpyxl.load_workbook(path, keep_vba=keep_vba, data_only=False)

    # 2) workbook de citire valori (valorile calculate din formule)
    wb_values = openpyxl.load_workbook(path, keep_vba=keep_vba, data_only=True)

    if "Sheet1" not in wb.sheetnames:
        raise RuntimeError("Nu găsesc Sheet1 în fișier.")
    if "Sheet1" not in wb_values.sheetnames:
        raise RuntimeError("Nu găsesc Sheet1 în fișier (data_only).")

    ws = wb["Sheet1"]
    ws_values = wb_values["Sheet1"]

    CopiereAutomataCombinata(ws)
    ActualizareColoana23(ws)
    AdaugaCategoryDinColoana21(ws)

    constant = build_constant_fast(wb, ws, ws_values)

    AdaugaEventNote(constant)
    UnireColoane119Si121(constant)
    ActualizareColoana123(constant)

    clear_sheet1_cols(ws)

    out = path.with_name(path.stem + "_modificat" + path.suffix)
    wb.save(out)
    return out

# =========================
# TELEGRAM
# =========================
@dp.message(CommandStart())
async def start(msg: types.Message):
    await msg.answer(
        "✅ Sunt online.\n"
        "Trimite-mi un fișier Excel (.xlsx sau .xlsm) și îl procesez.\n"
        "⚠️ .xls nu este suportat aici — salvează-l ca .xlsx."
    )

@dp.message(F.document)
async def handle_file(msg: types.Message):
    doc = msg.document
    filename = doc.file_name or "fisier.xlsx"
    ext = Path(filename).suffix.lower()

    if ext == ".xls":
        await msg.answer(
            "❌ Fișier .xls (format vechi) nu e suportat pe Render Free.\n"
            "✅ Te rog: Excel → Save As → .xlsx și retrimite."
        )
        return

    if ext not in ALLOWED_EXT:
        await msg.answer("❌ Format neacceptat. Trimite doar .xlsx sau .xlsm.")
        return

    await msg.answer("📥 Am primit fișierul. Îl descarc...")

    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        in_file = tmpdir / filename

        file = await bot.get_file(doc.file_id)
        await bot.download_file(file.file_path, in_file)

        await msg.answer("⚙️ Procesez fișierul ...")
        try:
            out_file = process_excel(in_file)
        except Exception as e:
            await msg.answer(f"❌ Eroare la procesare: {e}")
            return

        await msg.answer_document(types.FSInputFile(out_file), caption="✅ Gata. Iată fișierul procesat.")

# =========================
# HEALTH SERVER (Render Web Service)
# =========================
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
    await run_health_server()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
