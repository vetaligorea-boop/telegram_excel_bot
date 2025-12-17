import asyncio
import os
import tempfile
import re
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

ALLOWED_EXT = {".xlsx", ".xlsm"}   # .xls nu e suportat fara LibreOffice
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
    Converteste valoarea Excel (time) in HH:MM:SS.
    Accepta:
    - datetime / time -> ok
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
# FORMULA TIME EVALUATOR (pentru "=xx+xx")
# =========================
CELL_REF_RE = re.compile(r"^([A-Z]{1,3})(\d+)$")

def col_letters_to_index(letters: str) -> int:
    """A->1, B->2, ..., Z->26, AA->27 ..."""
    letters = letters.upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def parse_cell_ref(token: str):
    token = token.replace("$", "").strip().upper()
    m = CELL_REF_RE.match(token)
    if not m:
        return None
    col_letters, row_str = m.group(1), m.group(2)
    return int(row_str), col_letters_to_index(col_letters)

def excel_time_to_seconds(v) -> float:
    """Convertește orice reprezentare time într-un număr de secunde."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, datetime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, dtime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, (int, float)):
        # fracție de zi
        frac = float(v) % 1.0
        return frac * 86400.0
    # text "HH:MM:SS"
    s = safe_str(v).strip()
    if re.match(r"^\d{1,2}:\d{2}:\d{2}$", s):
        hh, mm, ss = s.split(":")
        return int(hh) * 3600 + int(mm) * 60 + int(ss)
    if s == "0":
        return 0.0
    return 0.0

def seconds_to_excel_fraction(seconds: float) -> float:
    return (seconds % 86400.0) / 86400.0

def eval_simple_time_formula(formula: str, ws, ws_values):
    """
    Evaluează formule SIMPLE cu + și -:
      =B5+E5
      =A1+C1
      =B2+0.001
      =B2- C2
    Nu suportă funcții (TIME(), IF(), etc.) – dacă există, return None.
    """
    if not isinstance(formula, str):
        return None
    f = formula.strip()
    if not f.startswith("="):
        return None

    expr = f[1:].replace(" ", "")

    # dacă are funcții, paranteze, etc. -> nu încercăm
    if "(" in expr or ")" in expr:
        return None

    # tokenizare simplă pe + și -
    # păstrăm operatorii
    parts = re.split(r"([+\-])", expr)
    if not parts:
        return None

    total_sec = None
    op = "+"

    for part in parts:
        part = part.strip()
        if part == "":
            continue
        if part in ("+", "-"):
            op = part
            continue

        sec = None

        # 1) cell ref?
        ref = parse_cell_ref(part)
        if ref:
            r, c = ref
            # preferăm valoarea "data_only" dacă există
            v = ws_values.cell(r, c).value
            if v in (None, ""):
                v = ws.cell(r, c).value
            sec = excel_time_to_seconds(v)

        else:
            # 2) numeric literal?
            try:
                num = float(part)
                # Excel numeric time fraction -> sec
                sec = excel_time_to_seconds(num)
            except Exception:
                # 3) text HH:MM:SS?
                sec = excel_time_to_seconds(part)

        if sec is None:
            return None

        if total_sec is None:
            total_sec = sec
        else:
            total_sec = (total_sec + sec) if op == "+" else (total_sec - sec)

    if total_sec is None:
        return None

    return seconds_to_excel_fraction(total_sec)


def get_time_value_for_row(ws, ws_values, row: int):
    """
    Ia ora pentru randul 'row' din Sheet1 col 2:
    - dacă data_only are valoare, o folosim
    - altfel, dacă e formulă, o calculăm noi (simple + / -)
    - altfel, luăm direct valoarea
    """
    # 1) încearcă valoarea calculată (cached) din data_only
    v_cached = ws_values.cell(row, 2).value
    if v_cached not in (None, "", 0, "0"):
        return v_cached

    # 2) verifică formula în workbook normal
    v_raw = ws.cell(row, 2).value
    if isinstance(v_raw, str) and v_raw.strip().startswith("="):
        v_eval = eval_simple_time_formula(v_raw, ws, ws_values)
        if v_eval is not None:
            return v_eval

    # 3) fallback direct
    return v_raw


# =========================
# LOGIC (partea ta)
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
# constant sheet build
# =========================
def build_constant_fast(wb, ws, ws_values):
    if "constant" in wb.sheetnames:
        constant = wb["constant"]
        if constant.max_row > 0:
            constant.delete_rows(1, constant.max_row)
    else:
        constant = wb.create_sheet("constant")

    # muta constant la final
    sheets = wb._sheets
    if constant in sheets:
        sheets.remove(constant)
        sheets.append(constant)

    lastRow = max(get_last_row_in_col(ws, 6), get_last_row_in_col(ws, 4))

    out_row = 1
    insidePlaylist = False
    firstValueFound = False

    def write_row(r19, r20, r21, r23, ora_value):
        nonlocal out_row

        set_cell(constant, out_row, 119, safe_str(r19).strip())
        set_cell(constant, out_row, 120, safe_str(r20))
        set_cell(constant, out_row, 121, safe_str(r21))
        set_cell(constant, out_row, 123, safe_str(r23))

        # ✅ ORA corecta
        set_cell(constant, out_row, 122, format_time_value(ora_value))

        v119 = safe_str(constant.cell(out_row, 119).value)
        v121 = safe_str(constant.cell(out_row, 121).value)
        set_cell(constant, out_row, 125, f"{v119} {v121}".strip())

        out_row += 1

    for i in range(1, lastRow + 1):
        col6 = safe_str(ws.cell(i, 6).value)

        # ✅ ora din Sheet1 col2 (inclusiv formule)
        ora_value = get_time_value_for_row(ws, ws_values, i)

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

    return constant


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

    # workbook principal (pastram VBA)
    wb = openpyxl.load_workbook(path, keep_vba=keep_vba, data_only=False)
    # workbook pentru values (daca exista cached results)
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

    build_constant_fast(wb, ws, ws_values)

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
            "❌ Fișier .xls (format vechi) nu e suportat aici.\n"
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
