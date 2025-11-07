import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

YELLOW = "FFFF00"
GREEN = "00B050"

font_g = Font(name="Arial", size=12, bold=True)
font_h = Font(name="Arial", size=12, bold=True)
font_i = Font(name="Arial", size=14, bold=True)
font_j = Font(name="Arial", size=14, bold=True)

fill_yellow = PatternFill(start_color=YELLOW, end_color=YELLOW, fill_type="solid")
fill_green = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")

align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_center = Alignment(horizontal="center", vertical="center")

thin = Side(style="thin")
thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
no_border = Border()


def get_last_row_in_column(ws, col_idx: int, start_row: int = 2) -> int:
    last = ws.max_row
    while last >= start_row:
        val = ws.cell(row=last, column=col_idx).value
        if val is not None and str(val).strip() != "":
            return last
        last -= 1
    return start_row


def process_workbook(input_path: str) -> str:
    file_name = os.path.basename(input_path)
    base, ext = os.path.splitext(file_name)

    if ext.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError("Format neacceptat (folosește .xlsx sau .xlsm)")

    keep_vba = ext.lower() == ".xlsm"
    wb = load_workbook(input_path, keep_vba=keep_vba)
    ws = wb.worksheets[0]

    last_row = get_last_row_in_column(ws, 7, 2)
    if last_row <= 1:
        wb.close()
        raise ValueError("Nu s-au găsit date în coloana G")

    for row in range(2, last_row + 1):
        c7 = ws.cell(row=row, column=7)
        c8 = ws.cell(row=row, column=8)
        c9 = ws.cell(row=row, column=9)
        c10 = ws.cell(row=row, column=10)
        c11 = ws.cell(row=row, column=11)

        # G
        c7.font = font_g
        c7.fill = fill_yellow
        c7.alignment = align_left

        # H
        val_h = c8.value
        if isinstance(val_h, (int, float)) and val_h > 0:
            total = int(val_h)
            h = total // 3600
            m = (total % 3600) // 60
            s = total % 60
            c8.value = f"{h:02d}:{m:02d}:{s:02d}"
        c8.font = font_h
        c8.fill = fill_yellow
        c8.alignment = align_right

        # I
        val_i = (c9.value if c9.value is not None else "").strip() if isinstance(c9.value, str) else c9.value
        if val_i not in (None, ""):
            c9.font = font_i
            c9.fill = fill_green
            c9.alignment = align_center
            c9.border = thin_border
        else:
            c9.border = no_border

        # J
        val_j_raw = c10.value
        if val_j_raw is not None and str(val_j_raw).strip() != "":
            s_val = str(val_j_raw).strip()
            if s_val.isdigit() and 1 <= int(s_val) <= 9:
                c10.value = f"_{s_val}__"
            else:
                c10.value = f"_{s_val}_"
        c10.font = font_j
        c10.fill = fill_yellow
        c10.alignment = align_center

        # K
        val_k = (c11.value if c11.value is not None else "").strip() if isinstance(c11.value, str) else c11.value
        if val_k not in (None, ""):
            c11.border = thin_border
        else:
            c11.border = no_border

    output_dir = os.path.dirname(input_path)
    new_name = f"{base}_modificat{ext}"
    output_path = os.path.join(output_dir, new_name)
    wb.save(output_path)
    wb.close()
    return output_path
