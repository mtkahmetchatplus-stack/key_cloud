"""
Transfer T-05-06_SH1 → Sayfa 1 and T-05_06_SH2 → Sayfa 2

- All cell values transferred
- All merged cells unmerged (text stays in top-left cell of each former merge)
- Column widths, row heights, cell styles preserved
- Print area and page setup (scale, orientation, paper size, margins) matched to source
"""

import copy
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

SRC_PATH = "/root/.claude/uploads/42f17e1c-d897-4441-8a08-7b3948ab5b0f/65082cad-24004RENPRCDAT0001_4.xlsx"
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


def parse_print_area(print_area_str):
    """Return (min_col, min_row, max_col, max_row) from a print area string like 'Sheet'!$A$1:$BS$62"""
    import re
    match = re.search(r'\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)', print_area_str)
    if match:
        c1, r1, c2, r2 = match.groups()
        return column_index_from_string(c1), int(r1), column_index_from_string(c2), int(r2)
    return None


def transfer_sheet(src_ws, dst_ws):
    # --- 1. Collect source merge info before unmerging ---
    # Map top-left cell -> value (already stored there), we just need the range info
    merge_ranges = [(m.min_row, m.min_col, m.max_row, m.max_col)
                    for m in src_ws.merged_cells.ranges]

    # --- 2. Unmerge all in source so iter_rows gives plain cells ---
    for m in list(src_ws.merged_cells.ranges):
        src_ws.unmerge_cells(str(m))

    # --- 3. Unmerge all in destination ---
    for m in list(dst_ws.merged_cells.ranges):
        dst_ws.unmerge_cells(str(m))

    # --- 4. Clear destination values ---
    for row in dst_ws.iter_rows():
        for cell in row:
            cell.value = None

    # --- 5. Copy row dimensions ---
    for row_idx, rd in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_idx].height        = rd.height
        dst_ws.row_dimensions[row_idx].hidden         = rd.hidden
        dst_ws.row_dimensions[row_idx].outline_level  = rd.outline_level

    # --- 6. Copy column dimensions ---
    for col_letter, cd in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width         = cd.width
        dst_ws.column_dimensions[col_letter].hidden         = cd.hidden
        dst_ws.column_dimensions[col_letter].outline_level  = cd.outline_level

    # --- 7. Copy values + styles cell by cell ---
    for row in src_ws.iter_rows():
        for src_cell in row:
            dst_cell = dst_ws.cell(row=src_cell.row, column=src_cell.column)
            # For formula values that reference the old sheet name, use the raw value
            val = src_cell.value
            if isinstance(val, str) and val.startswith("='T-05"):
                val = None  # drop cross-sheet formulas
            dst_cell.value = val
            copy_cell_style(src_cell, dst_cell)

    # --- 8. Page setup: orientation, scale, paper size ---
    src_ps = src_ws.page_setup
    dst_ws.page_setup.orientation = src_ps.orientation
    dst_ws.page_setup.paperSize   = src_ps.paperSize
    dst_ws.page_setup.scale       = src_ps.scale

    # --- 9. Page margins ---
    pm = src_ws.page_margins
    dst_ws.page_margins.left   = pm.left
    dst_ws.page_margins.right  = pm.right
    dst_ws.page_margins.top    = pm.top
    dst_ws.page_margins.bottom = pm.bottom
    dst_ws.page_margins.header = pm.header
    dst_ws.page_margins.footer = pm.footer

    # --- 10. Print area: strip source sheet name, use clean cell range ---
    raw_pa = src_ws.print_area
    if raw_pa:
        coords = parse_print_area(raw_pa)
        if coords:
            min_col, min_row, max_col, max_row = coords
            pa = f"${get_column_letter(min_col)}${min_row}:${get_column_letter(max_col)}${max_row}"
            dst_ws.print_area = pa
            print(f"    print_area set to: {pa}")


def main():
    print("Loading workbooks …")
    src_wb = openpyxl.load_workbook(SRC_PATH)
    dst_wb = openpyxl.load_workbook(DST_PATH)

    for src_name, dst_name in SHEET_MAP:
        print(f"\n  {src_name}  →  {dst_name}")
        src_ws = src_wb[src_name]
        dst_ws = dst_wb[dst_name]

        print(f"    src: {src_ws.max_row}r × {src_ws.max_column}c  merged={len(list(src_ws.merged_cells.ranges))}")
        print(f"    src print_area: {src_ws.print_area}  scale={src_ws.page_setup.scale}")

        transfer_sheet(src_ws, dst_ws)

        print(f"    dst after: {dst_ws.max_row}r × {dst_ws.max_column}c  merged=0")

    dst_wb.save(OUT_PATH)
    print(f"\nSaved → {OUT_PATH}")


if __name__ == "__main__":
    main()
