"""
Transfer T-05-06_SH1 → Sayfa 1 and T-05_06_SH2 → Sayfa 2
Source: Copy_of_24004RENPRCDAT0001_4.xlsx (cells already unmerged, values in place)
Target: KEY199MECDAT0005_AA.xlsx

Copies values, cell styles, column widths, row heights, page setup, margins, print area.
"""

import copy
import re
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

SRC_PATH = "/root/.claude/uploads/352a98e8-8adc-4e78-8f78-fbdee5caaa36/656d096f-Copy_of_24004RENPRCDAT0001_4.xlsx"
DST_PATH = "/root/.claude/uploads/42f17e1c-d897-4441-8a08-7b3948ab5b0f/4c49f826-KEY199MECDAT0005_AA.xlsx"
OUT_PATH = "/home/user/key_cloud/KEY199MECDAT0005_AA_updated.xlsx"

SHEET_MAP = [
    ("T-05-06_SH1", "Sayfa 1"),
    ("T-05_06_SH2", "Sayfa 2"),
]


def copy_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font          = copy.copy(src_cell.font)
        dst_cell.border        = copy.copy(src_cell.border)
        dst_cell.fill          = copy.copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection    = copy.copy(src_cell.protection)
        dst_cell.alignment     = copy.copy(src_cell.alignment)


def clean_print_area(raw):
    """Strip sheet name prefix and return clean $COL$ROW:$COL$ROW string."""
    match = re.search(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', raw)
    if match:
        c1, r1, c2, r2 = match.groups()
        return f"${c1}${r1}:${c2}${r2}"
    return raw


def transfer_sheet(src_ws, dst_ws):
    # 1. Unmerge destination (clean slate)
    for m in list(dst_ws.merged_cells.ranges):
        dst_ws.unmerge_cells(str(m))

    # 2. Clear destination values
    for row in dst_ws.iter_rows():
        for cell in row:
            if cell.__class__.__name__ != "MergedCell":
                cell.value = None

    # 3. Row dimensions
    for row_idx, rd in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_idx].height       = rd.height
        dst_ws.row_dimensions[row_idx].hidden        = rd.hidden
        dst_ws.row_dimensions[row_idx].outline_level = rd.outline_level

    # 4. Column dimensions
    for col_letter, cd in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width        = cd.width
        dst_ws.column_dimensions[col_letter].hidden        = cd.hidden
        dst_ws.column_dimensions[col_letter].outline_level = cd.outline_level

    # 5. Cell values + styles
    for row in src_ws.iter_rows():
        for src_cell in row:
            dst_cell = dst_ws.cell(row=src_cell.row, column=src_cell.column)
            dst_cell.value = src_cell.value
            copy_cell_style(src_cell, dst_cell)

    # 6. Replicate merged cells from source (source already has 0, but future-proof)
    for m in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(m))

    # 7. Page setup
    src_ps = src_ws.page_setup
    dst_ws.page_setup.orientation = src_ps.orientation
    dst_ws.page_setup.paperSize   = src_ps.paperSize
    dst_ws.page_setup.scale       = src_ps.scale

    # 8. Page margins
    pm = src_ws.page_margins
    dst_ws.page_margins.left   = pm.left
    dst_ws.page_margins.right  = pm.right
    dst_ws.page_margins.top    = pm.top
    dst_ws.page_margins.bottom = pm.bottom
    dst_ws.page_margins.header = pm.header
    dst_ws.page_margins.footer = pm.footer

    # 9. Print area (strip source sheet name)
    if src_ws.print_area:
        dst_ws.print_area = clean_print_area(src_ws.print_area)


def main():
    print("Loading workbooks …")
    src_wb = openpyxl.load_workbook(SRC_PATH)
    dst_wb = openpyxl.load_workbook(DST_PATH)

    for src_name, dst_name in SHEET_MAP:
        src_ws = src_wb[src_name]
        dst_ws = dst_wb[dst_name]

        print(f"\n  {src_name}  →  {dst_name}")
        print(f"    src: {src_ws.max_row}r × {src_ws.max_column}c  merged={len(list(src_ws.merged_cells.ranges))}")
        print(f"    print_area={src_ws.print_area}  scale={src_ws.page_setup.scale}  orientation={src_ws.page_setup.orientation}")

        transfer_sheet(src_ws, dst_ws)

        vals = sum(1 for row in dst_ws.iter_rows() for c in row if c.value is not None)
        print(f"    dst: {dst_ws.max_row}r × {dst_ws.max_column}c  merged={len(list(dst_ws.merged_cells.ranges))}  values={vals}")
        print(f"    dst print_area={dst_ws.print_area}")

    dst_wb.save(OUT_PATH)
    print(f"\nSaved → {OUT_PATH}")


if __name__ == "__main__":
    main()
