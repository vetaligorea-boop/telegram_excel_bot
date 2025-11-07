import os
from datetime import datetime, time
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Culori si stiluri
YELLOW = "FFFF00"
GREEN = "00B050"
RED = "FF0000"
LIGHT_GREEN = "00FF00"

font12_bold = Font(name="Arial", size=12, bold=True)
font14_bold = Font(name="Arial", size=14, bold=True)

fill_yellow = PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")
fill_green = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
fill_red = PatternFill(start_color=RED, end_color=RED, fill_type="solid")
fill_light_green = PatternFill(start_color=LIGHT_GREEN, end_color=LIGHT_GREEN, fill_type="solid")

align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_center = Alignment(horizontal="center", vertical="center")

thin = Side(style="thin")
thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
no_border = Border()


# ================= HELPERI GENERALI =================

def nz(val) -> str:
    if val is None:
        return ""
    try:
        return str(val)
    except Exception:
        return ""


def get_last_row(ws, col_idx: int) -> int:
    last = ws.max_row
    while last > 0:
        v = ws.cell(row=last, column=col_idx).value
        if v not in (None, "", " "):
            return last
        last -= 1
    return 1


def parse_excel_time(cell):
    v = cell.value
    if v is None:
        return None
    # deja datetime/time
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, time):
        return v
    # ca text
    try:
        txt = str(v).strip()
        if not txt:
            return None
        # încearcă HH:MM sau HH:MM:SS
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(txt, fmt).time()
            except ValueError:
                pass
    except Exception:
        return None
    return None


def in_interval(t: time, start_s: str, end_s: str) -> bool:
    try:
        s = datetime.strptime(start_s, "%H:%M:%S").time()
        e = datetime.strptime(end_s, "%H:%M:%S").time()
    except ValueError:
        return False
    return s <= t <= e


def file_format_from_ext(ext: str) -> int:
    ext = ext.lower()
    if ext == ".xlsx":
        return 51
    if ext == ".xlsm":
        return 52
    if ext == ".xls":
        return 56
    return 51


# ================= 1) FORMAT PUB_Zero (coloane G-K) =================

def format_pub_zero(pub_zero_path: str) -> str:
    """
    Echivalent ConvertAndFormatCols7_8_9_10_11 pentru un singur fisier.
    Intoarce calea noului fisier *_modificat.*
    """
    if not os.path.isfile(pub_zero_path):
        raise ValueError("Fisierul PUB_Zero nu exista.")

    base = os.path.basename(pub_zero_path)
    name, ext = os.path.splitext(base)
    if ext.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("PUB_Zero trebuie sa fie .xlsx sau .xlsm")

    keep_vba = ext.lower() == ".xlsm"
    wb = load_workbook(pub_zero_path, keep_vba=keep_vba)
    if not wb.worksheets:
        wb.close()
        raise ValueError("PUB_Zero nu contine foi.")
    ws = wb.worksheets[0]

    last_row = get_last_row(ws, 7)
    if last_row <= 1:
        wb.close()
        raise ValueError("PUB_Zero nu are date in coloana G.")

    for r in range(2, last_row + 1):
        c7 = ws.cell(r, 7)
        c8 = ws.cell(r, 8)
        c9 = ws.cell(r, 9)
        c10 = ws.cell(r, 10)
        c11 = ws.cell(r, 11)

        # G
        c7.font = font12_bold
        c7.fill = fill_yellow
        c7.alignment = align_left

        # H
        v = c8.value
        if isinstance(v, (int, float)) and v > 0:
            total = int(v)
            h = total // 3600
            m = (total % 3600) // 60
            s = total % 60
            c8.value = f"{h:02d}:{m:02d}:{s:02d}"
        c8.font = font12_bold
        c8.fill = fill_yellow
        c8.alignment = align_right

        # I
        if c9.value not in (None, "", " "):
            c9.font = font14_bold
            c9.fill = fill_green
            c9.alignment = align_center
            c9.border = thin_border
        else:
            c9.border = no_border

        # J
        if c10.value not in (None, "", " "):
            s_val = str(c10.value).strip()
            if s_val.isdigit() and 1 <= int(s_val) <= 9:
                c10.value = f"_{s_val}__"
            else:
                c10.value = f"_{s_val}_"
        c10.font = font14_bold
        c10.fill = fill_yellow
        c10.alignment = align_center

        # K
        if c11.value not in (None, "", " "):
            c11.border = thin_border
        else:
            c11.border = no_border

    out_dir = os.path.dirname(pub_zero_path)
    new_name = f"{name}_modificat{ext}"
    out_path = os.path.join(out_dir, new_name)
    wb.save(out_path)
    wb.close()
    return out_path


# ================= 2) FLOW COMBINAT (IN + PUB_IN -> FINAL) =================

EXCLUDE_EXACT = {
    "id_jtv_2024_dua_lipa_dance_the_night",
    "id_jtv_2024_miley_cyrus_flowers",
    "id_jtv_2024_the weeknd_ariana grande_save_your_tears",
    "id 15 ani_25sec_v1",
    "youtube sofia obada jurnalul orei 19 ok",
    "jurnalsportiv",
    "meteonew",
}


def col4_este_exclus(val: str) -> bool:
    if val is None:
        return False
    v = nz(val).strip().lower()
    if v.startswith("id pub") or v.startswith("id_pub_") or v.startswith("id promo") or v.startswith("id_promo_"):
        return True
    if v.startswith("interzis") or v.startswith("cca_") or v.startswith("cca orele"):
        return True
    if v in EXCLUDE_EXACT:
        return True
    return False


def colorare_rosu_col_e(ws):
    last_row = max(get_last_row(ws, 4), get_last_row(ws, 5))
    if last_row < 1:
        return
    for r in range(1, last_row + 1):
        val_d = nz(ws.cell(r, 4).value)
        val_e = nz(ws.cell(r, 5).value)
        if val_e != "" and not col4_este_exclus(val_d):
            c = ws.cell(r, 5)
            c.fill = fill_red


def sterge_intre_playlist(ws):
    col_f = 6
    last_row = get_last_row(ws, col_f)
    i = 1
    while i <= last_row:
        val = nz(ws.cell(i, col_f).value)
        if val.startswith("PLAYLIST_IN_"):
            start_row = i
            end_row = 0
            j = i + 1
            while j <= last_row:
                v2 = nz(ws.cell(j, col_f).value)
                if v2.startswith("PLAYLIST_OUT_"):
                    end_row = j
                    break
                j += 1
            if end_row > 0:
                if end_row > start_row + 1:
                    ws.delete_rows(start_row + 1, end_row - start_row - 1)
                    last_row = get_last_row(ws, col_f)
                i = start_row
            else:
                break
        i += 1


def proceseaza_interval(ws_in, ws_out, ora_start, ora_end, variante):
    # colectam iduri + G,H,I din wsIN pentru interval
    last_in = get_last_row(ws_in, 3)
    iduri = []
    colG = []
    colH = []
    colI = []

    r = 1
    while r <= last_in:
        t = parse_excel_time(ws_in.cell(r, 3))
        if t and in_interval(t, ora_start, ora_end):
            rr = r
            while rr <= last_in:
                t2 = parse_excel_time(ws_in.cell(rr, 3))
                if t2 and not in_interval(t2, ora_start, ora_end):
                    break
                if nz(ws_in.cell(rr, 10).value) != "":
                    iduri.append(nz(ws_in.cell(rr, 10).value))
                    colG.append(nz(ws_in.cell(rr, 7).value))
                    colH.append(nz(ws_in.cell(rr, 8).value))
                    colI.append(nz(ws_in.cell(rr, 9).value))
                rr += 1
            break
        r += 1

    if not iduri:
        return

    last_out = get_last_row(ws_out, 6)

    for var in variante:
        start_row = 0
        end_row = 0
        for r in range(1, last_out + 1):
            txt = nz(ws_out.cell(r, 6).value)
            if txt == f"PLAYLIST_IN_{var}":
                start_row = r
            if txt == f"PLAYLIST_OUT_{var}":
                end_row = r
                break

        if start_row > 0 and end_row > 0 and end_row > start_row:
            # sterge interior
            if end_row > start_row + 1:
                ws_out.delete_rows(start_row + 1, end_row - start_row - 1)
                end_row = start_row + 1

            # insereaza iduri
            ws_out.insert_rows(end_row, len(iduri))
            for i, _ in enumerate(iduri, start=1):
                rdest = start_row + i
                ws_out.cell(rdest, 4, colG[i-1])
                ws_out.cell(rdest, 5, colH[i-1])
                ws_out.cell(rdest, 6, iduri[i-1])
                ws_out.cell(rdest, 7, colI[i-1])

                c4 = ws_out.cell(rdest, 4)
                c5 = ws_out.cell(rdest, 5)
                c6 = ws_out.cell(rdest, 6)
                c7 = ws_out.cell(rdest, 7)

                for c, align in ((c4, align_left), (c5, align_right), (c6, align_center)):
                    c.font = font14_bold
                    c.alignment = align
                    c.fill = fill_yellow
                if nz(c7.value) != "":
                    c7.fill = fill_light_green
                    c7.border = thin_border
            break  # doar prima varianta valida


def flow_combinat(in_path: str, pub_in_path: str) -> str:
    """
    Echivalent Proceseaza_Flow_Combinat simplificat:
    - ia fisierul IN (playlist baza)
    - coloreaza E in rosu cu excluderi
    - sterge randuri intre PLAYLIST_IN_/PLAYLIST_OUT_
    - ia primul sheet din PUB_IN
    - insereaza intervale dupa reguli
    - salveaza FINAL_modificat.*
    """
    if not os.path.isfile(in_path):
        raise ValueError("Fisierul IN nu exista.")
    if not os.path.isfile(pub_in_path):
        raise ValueError("Fisierul PUB_IN nu exista.")

    in_base = os.path.basename(in_path)
    name_in, ext_in = os.path.splitext(in_base)
    keep_vba = ext_in.lower() == ".xlsm"

    wb_out = load_workbook(in_path, keep_vba=keep_vba)
    if not wb_out.worksheets:
        wb_out.close()
        raise ValueError("Fisierul IN nu contine foi.")
    ws_out = wb_out.worksheets[0]

    # pas 1: rosu pe col E
    colorare_rosu_col_e(ws_out)

    # pas 2a: sterge intre PLAYLIST_IN_/OUT_
    sterge_intre_playlist(ws_out)

    # pas 2b: ia date din PUB_IN
    wb_in = load_workbook(pub_in_path, keep_vba=False)
    if not wb_in.worksheets:
        wb_in.close()
        wb_out.close()
        raise ValueError("Fisierul PUB_IN nu contine foi.")
    ws_in = wb_in.worksheets[0]

    # definim intervalele ca in macro
    ore_def = [
        ("06:00:00","06:30:00",["06_30","06_20","06_10"]),
        ("06:30:01","06:59:00",["06_50","06_40","06_45"]),
        ("07:00:00","07:30:00",["07_20","07_10","07_30"]),
        ("07:31:00","07:59:00",["07_50","07_40","07_45"]),
        ("08:00:00","08:31:00",["08_20","08_10","08_30"]),
        ("08:32:00","08:59:00",["08_50","08_40","08_45"]),
        ("09:00:00","09:31:00",["09_20","09_10","09_30"]),
        ("09:32:00","09:59:00",["09_50","09_40","09_45"]),
        ("10:00:00","10:31:00",["10_20","10_10","10_30"]),
        ("10:32:00","10:59:00",["10_50","10_40","10_45"]),
        ("11:00:00","11:31:00",["11_20","11_10","11_30"]),
        ("11:32:00","11:59:00",["11_50","11_40","11_45"]),
        ("12:00:00","12:31:00",["12_20","12_10","12_30"]),
        ("12:32:00","12:59:00",["12_50","12_40","12_45"]),
        ("13:00:00","13:31:00",["13_20","13_10","13_30"]),
        ("13:32:00","13:59:00",["13_50","13_40","13_45"]),
        ("14:00:00","14:31:00",["14_20","14_10","14_30"]),
        ("14:32:00","14:59:00",["14_50","14_40","14_45"]),
        ("15:00:00","15:31:00",["15_20","15_10","15_30"]),
        ("15:32:00","15:59:00",["15_50","15_40","15_45"]),
        ("16:00:00","16:31:00",["16_20","16_10","16_30"]),
        ("16:32:00","16:59:00",["16_50","16_40","16_45"]),
        ("17:00:00","17:31:00",["17_20","17_10","17_30"]),
        ("17:32:00","17:59:00",["17_50","17_40","17_45"]),
        ("18:00:00","18:31:00",["18_20","18_10","18_30"]),
        ("18:32:00","18:59:00",["18_50","18_40","18_45"]),
        ("19:00:00","19:31:00",["19_20","19_10","19_30"]),
        ("19:32:00","19:59:00",["19_50","19_40","19_45"]),
        ("20:00:00","20:31:00",["20_20","20_10","20_30"]),
        ("20:32:00","20:59:00",["20_50","20_40","20_45"]),
        ("21:00:00","21:31:00",["21_20","21_10","21_30"]),
        ("21:32:00","21:59:00",["21_50","21_40","21_45"]),
        ("22:00:00","22:31:00",["22_20","22_10","22_30"]),
        ("22:32:00","22:59:00",["22_50","22_40","22_45"]),
        ("23:00:00","23:31:00",["23_20","23_10","23_30"]),
        ("23:32:00","23:59:00",["23_50","23_40","23_45"]),
        ("00:00:00","00:31:00",["00_20","00_10","00_30"]),
        ("00:32:00","00:59:00",["00_50","00_40","00_45"]),
        ("01:00:00","01:31:00",["01_20","01_10","01_30"]),
        ("01:32:00","01:59:00",["01_50","01_40","01_45"]),
    ]

    for start_s, end_s, vars_list in ore_def:
        proceseaza_interval(ws_in, ws_out, start_s, end_s, vars_list)

    wb_in.close()

    out_dir = os.path.dirname(in_path)
    new_name = f"{name_in}_modificat{ext_in}"
    final_path = os.path.join(out_dir, new_name)
    fmt = file_format_from_ext(ext_in)
    wb_out.save(final_path)
    wb_out.close()

    return final_path


# ================= 3) FUNCTIE PRINCIPALA PENTRU BOT =================

def process_pair(in_path: str, pub_zero_path: str) -> str:
    """
    Pentru bot:
    - primeste fisier IN si fisier PUB_Zero
    - genereaza PUB_IN din PUB_Zero (format G-K)
    - ruleaza flow combinat IN + PUB_IN
    - intoarce calea fisierului FINAL (_modificat)
    """
    pub_in_path = format_pub_zero(pub_zero_path)
    final_path = flow_combinat(in_path, pub_in_path)
    return final_path
