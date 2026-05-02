import os
import re
import openpyxl
from openpyxl import load_workbook
from copy import copy

INPUT_FILE = r'c:\Users\USER\Desktop\Coding Projects\Automated Quotations\excel_inputs_files\Standard Decora Door Prices Gefalzt.xlsx'
OUTPUT_DIR = r'c:\Users\USER\Desktop\Coding Projects\Automated Quotations\excel_output_files'

os.makedirs(OUTPUT_DIR, exist_ok=True)

wb_src = load_workbook(INPUT_FILE)
ws_src = wb_src.active


def safe_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '_', str(name)).strip()


def copy_cell(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.alignment = copy(src_cell.alignment)


header_row = list(ws_src.iter_rows(min_row=1, max_row=1))[0]

# Group data rows by door name (column B = Oberflache)
doors = {}
for row in ws_src.iter_rows(min_row=2, max_row=ws_src.max_row):
    door_name = row[1].value  # column B
    if door_name:
        if door_name not in doors:
            doors[door_name] = []
        doors[door_name].append(row)

print(f"Found {len(doors)} doors. Writing files...")

for door_name, rows in doors.items():
    wb_new = openpyxl.Workbook()
    ws_new = wb_new.active
    ws_new.title = "door_leaf_prices_full"

    # Copy header
    for col_idx, src_cell in enumerate(header_row, 1):
        dst_cell = ws_new.cell(row=1, column=col_idx)
        copy_cell(src_cell, dst_cell)

    # Copy door rows
    for dst_row_idx, src_row in enumerate(rows, 2):
        for col_idx, src_cell in enumerate(src_row, 1):
            dst_cell = ws_new.cell(row=dst_row_idx, column=col_idx)
            copy_cell(src_cell, dst_cell)

    # Copy column widths
    for col_letter, col_dim in ws_src.column_dimensions.items():
        ws_new.column_dimensions[col_letter].width = col_dim.width

    filename = safe_filename(door_name) + ".xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    wb_new.save(filepath)
    print(f"  Saved: {filename}")

print(f"\nDone. {len(doors)} files written to {OUTPUT_DIR}")
