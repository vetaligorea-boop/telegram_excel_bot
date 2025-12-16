import asyncio
import os
import shutil
import tempfile
from pathlib import Path
from datetime import datetime, time

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

ALLOWED_EXT = {".xlsx", ".xlsm"}  # .xls nu este suportat aici


WHITE_FONT = Font(color="FFFFFFFF")  # alb (ARGB)


# =========================
# HELPERS (Excel)
# =========================
def safe_str(v) -> str:
    return "" if v is None else str(v)

def is_text(v) -> bool:
    return isinstance(v, str)

def startswith_any(s: str, prefixes) -> bool:
    s2 = s or ""
    return any(s2.startswith(p) for p in prefixes)

def like_prefix(s: str, prefix: str) -> bool:
    # echivalent simplu pentru VBA Like "PLAYLIST_IN_*"
    return (s or "").startswith(prefix)

def get_last_row_in_col(ws, col: int) -> int:
    # caută ultimul rând cu ceva în col, de jos în sus
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, col).value not in (None, ""):
            return r
    return 1

def get_fill_rgb_safe(cell):
    try:
        f = cell.fill
        if not f or not f.patternType:
            return None
        c = f.fgColor
        return c.rgb
    except Exception:
        return None

def is_yellow_colorindex6(cell) -> bool:
    # în fișierul tău, galbenul apare de obicei ca FFFFFF00
    rgb = get_fill_rgb_safe(cell)
    return rgb in ("FFFFFF00", "FFFF00")


def set_cell(ws, row, col, value):
    c = ws.cell(row, col)
    c.value = value
    c.font = WHITE_FONT


# =========================
# VBA -> PYTHON TRANSLATION
# =========================

def CopiereAutomataCombinata(ws):
    # lastRow după coloana D (4)
    lastRow = get_last_row_in_col(ws, 4)
    insidePlaylist = False

    for i in range(1, lastRow + 1):
        cell_d = ws.cell(i, 4)
        if not is_yellow_colorindex6(cell_d):  # exclude galben
            cellValue = safe_str(cell_d.value)

            col6 = safe_str(ws.cell(i, 6).value)
            col3 = ws.cell(i, 3).value

            # ID PUB_* => col19 dacă NU e playlist marker
            if (cellValue.startswith("ID PUB_") or cellValue.startswith("ID_PUB_") or cellValue.startswith("ID PUB")):
                if not (like_prefix(col6, "PLAYLIST_IN_") or like_prefix(col6, "PLAYLIST_OUT_")):
                    set_cell(ws, i, 19, cellValue)

            # dacă col3 nu e gol => col21
            elif col3 not in (None, ""):
                set_cell(ws, i, 21, cellValue)

            # altfel => col19 (dacă col21 e gol)
            else:
                if safe_str(ws.cell(i, 21).value) == "":
                    set_cell(ws, i, 19, cellValue)

        # Copiere între PLAYLIST_IN_* și PLAYLIST_OUT_* din col6 în col20
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

        # CCA doar în col20
        dval = safe_str(ws.cell(i, 4).value)
        if dval in ("CCA_SARE_ZAHAR_GRASIMI", "CCA_ORELE_MESEI", "CCA_GIMNASTICA", "CCA_FRUCTE", "CCA_BEA_APA"):
            set_cell(ws, i, 20, dval)
            ws.cell(i, 19).value = None

        # Pentru rândurile cu PLAYLIST_IN/OUT => col20 = col4, col19 gol
        col6_val2 = safe_str(ws.cell(i, 6).value)
        if like_prefix(col6_val2, "PLAYLIST_IN_") or like_prefix(col6_val2, "PLAYLIST_OUT_"):
            set_cell(ws, i, 20, safe_str(ws.cell(i, 4).value))
            ws.cell(i, 19).value = None


def InsertCUB_PUB_TEST(ws):
    lastRow = get_last_row_in_col(ws, 6)
    insidePlaylist = False
    firstValueFound = False
    skipNext = False

    i = 1
    while i <= lastRow:
        col6 = safe_str(ws.cell(i, 6).value)

        if like_prefix(col6, "PLAYLIST_IN_"):
            insidePlaylist = True
            firstValueFound = False
            skipNext = False

        elif like_prefix(col6, "PLAYLIST_OUT_"):
            insidePlaylist = False

        elif insidePlaylist and not skipNext:
            # VBA folosea col6 ca „valoare”; aici păstrăm aceeași logică
            if not firstValueFound:
                firstValueFound = True
            else:
                ws.insert_rows(i)  # inserează deasupra rândului i (ca în VBA insert pe i)
                set_cell(ws, i, 20, "ID CUB_PUB_TEST")
                lastRow += 1
                skipNext = True  # sari peste rândul inserat

        else:
            skipNext = False

        i += 1


def AdaugaCubPubTestInColoana19(ws):
    excludeList = [
        "Jurnalfinanciar", "Jurnalulfinanciar", "PROMO YOUTUBE", "YOUTUBE SOFIA ",
        "JurnalSportivNEW", "JurnalulSportiv NEW", "JurnalSportiv ", "JurnalulSportiv ",
        "Carton ap", "Carton 12", "Carton 15", "INTERZIS_AP", "INTERZIS_12", "INTERZIS_15",
        "EarthTV", "stirile deficienta auz 17 HD", "MeteoNEW", "PostScriptumNEW",
        "Promourile", "bumper", "Studio comentarii", "Fotbal Repriza 1", "Stiri 10 min",
        "Fotbal Repriza 2", "REZUMATE", "DE FACTO Cioban", "DE FACTO Ciobanu",
        "PLANUL EUROPA CIOBANU", "PLANUL EUROPA CIOBAN", "DeFacto Tulbure", "PLANUL EUROPA Tulbure",
        "ID_Promo_", "INTERZIS_", "ID_PROMO_", "ID PROMO_", "ID_PROMO", "ID_Promo_",
        "ID_PUB_", "ID PUB_", "ID_PUB", "____PROMOURIIII___"
    ]

    lastRow = get_last_row_in_col(ws, 19)

    # de jos în sus (ca în VBA)
    for i in range(lastRow - 1, 0, -1):
        cellValue = safe_str(ws.cell(i, 19).value)
        nextCellValue = safe_str(ws.cell(i + 1, 19).value)
        col124Value = safe_str(ws.cell(i, 124).value)

        # dacă col124 conține "#EVENT NOTE" => skip
        if "#event note" in col124Value.lower():
            continue

        # exclude list check (case-insensitive, contains)
        excludeCurrent = False
        cv_low = cellValue.lower()
        nv_low = nextCellValue.lower()
        for ex in excludeList:
            ex_low = ex.lower()
            if ex_low in cv_low or ex_low in nv_low:
                excludeCurrent = True
                break

        if excludeCurrent:
            continue

        # prefix / suffix rules (simplificat ca în VBA)
        if cellValue == "":
            continue

        bad_prefix = (
            cellValue.startswith("ID PROMO_") or cellValue.startswith("ID_PROMO_") or cellValue.startswith("ID PROMO") or
            cellValue.startswith("ID PUB_") or cellValue.startswith("ID_PUB_") or cellValue.startswith("ID PUB")
        )

        if bad_prefix:
            continue

        # inserează rând sub i
        ws.insert_rows(i + 1)
        set_cell(ws, i + 1, 19, "ID CUB_PUB_TEST")


def ActualizareColoana23(ws):
    lastRow = get_last_row_in_col(ws, 6)
    insidePlaylist = False

    for i in range(1, lastRow + 1):
        col6 = safe_str(ws.cell(i, 6).value)

        if like_prefix(col6, "PLAYLIST_IN_"):
            insidePlaylist = True
            set_cell(ws, i, 23, "pub_start   #COLOR 65535")

        elif like_prefix(col6, "PLAYLIST_OUT_"):
            insidePlaylist = False
            set_cell(ws, i, 23, "pub_stop   #COLOR 4227327")

        else:
            # în VBA: se uită la valoarea existentă din col23 și o normalizează
            cellValue = safe_str(ws.cell(i, 23).value).strip().lower()

            if cellValue == "ceas":
                set_cell(ws, i, 23, "ceas   #COLOR 8421631")
            elif cellValue == "ap":
                set_cell(ws, i, 23, "ap   #COLOR 8454016")
            elif cellValue in ("cr+ap", "cr+12"):
                set_cell(ws, i, 23, f"{cellValue}   #COLOR 16777088")
            elif cellValue in ("reluare_in", "reluare_mid", "reluare_out"):
                set_cell(ws, i, 23, f"{cellValue}   #COLOR 16744448")
            elif cellValue in ("premiera_in", "premiera_mid", "premiera_out", "premiera_ap", "premiera_12", "premiera_15"):
                set_cell(ws, i, 23, f"{cellValue}   #COLOR 8388863")
            else:
                # păstrează valoarea originală (dar pune font alb dacă are ceva)
                if ws.cell(i, 23).value not in (None, ""):
                    ws.cell(i, 23).font = WHITE_FONT


def AdaugaCategoryDinColoana21(ws):
    excludeList = {"FILLER", "Ceas + Direct", "CEC"}
    lastRow = get_last_row_in_col(ws, 21)

    for i in range(1, lastRow + 1):
        col21 = ws.cell(i, 21).value
        col3 = safe_str(ws.cell(i, 3).value)

        if col21 not in (None, "") and col3 not in excludeList:
            # fără prefix #CATEGORY (ca în VBA)
            set_cell(ws, i, 23, col3)


def CreareSheetConstant(wb, ws_sheet1):
    # dacă nu există constant => creează
    if "constant" in wb.sheetnames:
        constantSheet = wb["constant"]
    else:
        constantSheet = wb.create_sheet("constant")

    # mută constant la final
    # (openpyxl: mută în lista internă)
    sheets = wb._sheets
    if constantSheet in sheets:
        sheets.remove(constantSheet)
        sheets.append(constantSheet)

    lastRow = get_last_row_in_col(ws_sheet1, 19)

    for i in range(1, lastRow + 1):
        cellValue19 = safe_str(ws_sheet1.cell(i, 19).value).strip()

        if cellValue19 == "____PROMOURIIII___":
            constantSheet.cell(i, 119).value = None
        else:
            set_cell(constantSheet, i, 119, cellValue19)

        set_cell(constantSheet, i, 120, safe_str(ws_sheet1.cell(i, 20).value))
        set_cell(constantSheet, i, 121, safe_str(ws_sheet1.cell(i, 21).value))
        set_cell(constantSheet, i, 123, safe_str(ws_sheet1.cell(i, 23).value))

        # copie col2 -> constant col122 (în VBA aici era direct)
        set_cell(constantSheet, i, 122, safe_str(ws_sheet1.cell(i, 2).value))

        # concat col119 + col121 -> col125
        set_cell(constantSheet, i, 125, f"{safe_str(constantSheet.cell(i, 119).value)} {safe_str(constantSheet.cell(i, 121).value)}".strip())

    return constantSheet


def format_time_value(v) -> str:
    # VBA: Format(ws.Cells(i,2).Value, "hh:nn:ss")
    if isinstance(v, datetime):
        return v.strftime("%H:%M:%S")
    if isinstance(v, time):
        return v.strftime("%H:%M:%S")
    # dacă vine ca număr Excel (float), openpyxl de obicei îl citește ca datetime/time, dar păstrăm fallback:
    return safe_str(v)


def AdaugaColoana2InColoana122(ws_sheet1, constantSheet):
    lastRow = get_last_row_in_col(constantSheet, 119)
    for i in range(1, lastRow + 1):
        t = format_time_value(ws_sheet1.cell(i, 2).value)
        set_cell(constantSheet, i, 122, t)


def AdaugaEventNote(constantSheet):
    lastRow = get_last_row_in_col(constantSheet, 119)

    # pentru col119: prima valoare din grup (până la blank)
    eventNoteAdded119 = False
    for i in range(1, lastRow + 1):
        v119 = safe_str(constantSheet.cell(i, 119).value)
        v122 = safe_str(constantSheet.cell(i, 122).value)

        if v119 != "" and not eventNoteAdded119:
            set_cell(constantSheet, i, 124, f"{v122} {v119}".strip())
            eventNoteAdded119 = True
        elif v119 == "":
            eventNoteAdded119 = False

    # pentru col120: prima valoare din grup (până la blank)
    eventNoteAdded120 = False
    for i in range(1, lastRow + 1):
        v120 = safe_str(constantSheet.cell(i, 120).value)
        v122 = safe_str(constantSheet.cell(i, 122).value)

        if v120 != "" and not eventNoteAdded120:
            set_cell(constantSheet, i, 124, f"{v122} {v120}".strip())
            eventNoteAdded120 = True
        elif v120 == "":
            eventNoteAdded120 = False

    # pentru col121: întotdeauna dacă există valoare
    for i in range(1, lastRow + 1):
        v121 = safe_str(constantSheet.cell(i, 121).value)
        v122 = safe_str(constantSheet.cell(i, 122).value)
        if v121 != "":
            set_cell(constantSheet, i, 124, f"{v122} {v121}".strip())


def UnireColoane119Si121(constantSheet):
    lastRow = get_last_row_in_col(constantSheet, 119)
    for i in range(1, lastRow + 1):
        v119 = safe_str(constantSheet.cell(i, 119).value)
        v121 = safe_str(constantSheet.cell(i, 121).value)
        set_cell(constantSheet, i, 125, f"{v119} {v121}".strip())


def ResetSheet1(ws):
    # șterge rândurile care conțin "ID CUB_PUB_TEST" în col 19/20/21/23
    lastRow = max(get_last_row_in_col(ws, 19), get_last_row_in_col(ws, 20), get_last_row_in_col(ws, 21), get_last_row_in_col(ws, 23))

    for i in range(lastRow, 0, -1):
        if (safe_str(ws.cell(i, 19).value) == "ID CUB_PUB_TEST" or
            safe_str(ws.cell(i, 20).value) == "ID CUB_PUB_TEST" or
            safe_str(ws.cell(i, 21).value) == "ID CUB_PUB_TEST" or
            safe_str(ws.cell(i, 23).value) == "ID CUB_PUB_TEST"):
            ws.delete_rows(i)

    # curăță coloanele 19,20,21,23
    lastRow2 = get_last_row_in_col(ws, 4)  # ca să curățăm până la zona reală
    for i in range(1, lastRow2 + 1):
        ws.cell(i, 19).value = None
        ws.cell(i, 20).value = None
        ws.cell(i, 21).value = None
        ws.cell(i, 23).value = None


def ActualizareColoana123(constantSheet):
    lastRow = get_last_row_in_col(constantSheet, 123)
    excludeList = {"ceas+direct", "ceas + direct", "."}

    for i in range(1, lastRow + 1):
        cellValue = safe_str(constantSheet.cell(i, 123).value).strip()
        low = cellValue.lower()

        if low in excludeList:
            constantSheet.cell(i, 123).value = None
            continue

        if low == "pub_start":
            set_cell(constantSheet, i, 123, "pub_start   #COLOR 65535")
        elif low == "pub_stop":
            set_cell(constantSheet, i, 123, "pub_stop   #COLOR 4227327")
        elif low == "ceas":
            set_cell(constantSheet, i, 123, "ceas   #COLOR 8421631")
        elif low == "ap":
            set_cell(constantSheet, i, 123, "ap   #COLOR 8454016")
        elif low in ("cr+ap", "cr+12"):
            set_cell(constantSheet, i, 123, f"{low}   #COLOR 16777088")
        elif low in ("reluare_in", "reluare_mid", "reluare_out"):
            set_cell(constantSheet, i, 123, f"{low}   #COLOR 16744448")
        elif low in ("premiera_in", "premiera_mid", "premiera_out", "premiera_ap", "premiera_12", "premiera_15"):
            set_cell(constantSheet, i, 123, f"{low}   #COLOR 8388863")
        else:
            if constantSheet.cell(i, 123).value not in (None, ""):
                constantSheet.cell(i, 123).font = WHITE_FONT


# =========================
# MAIN EXCEL PROCESS
# =========================
def process_excel(path: Path) -> Path:
    # păstrăm macro-urile dacă e .xlsm
    keep_vba = (path.suffix.lower() == ".xlsm")
    wb = openpyxl.load_workbook(path, keep_vba=keep_vba)

    if "Sheet1" not in wb.sheetnames:
        raise RuntimeError("Nu găsesc Sheet1 în fișier.")

    ws = wb["Sheet1"]

    # Ordinea EXACT ca VBA
    CopiereAutomataCombinata(ws)
    InsertCUB_PUB_TEST(ws)
    AdaugaCubPubTestInColoana19(ws)
    ActualizareColoana23(ws)
    AdaugaCategoryDinColoana21(ws)
    constantSheet = CreareSheetConstant(wb, ws)
    AdaugaColoana2InColoana122(ws, constantSheet)
    AdaugaEventNote(constantSheet)
    UnireColoane119Si121(constantSheet)
    ResetSheet1(ws)
    ActualizareColoana123(constantSheet)

    out = path.with_name(path.stem + "_modificat" + path.suffix)
    wb.save(out)
    return out


# =========================
# TELEGRAM HANDLERS
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
