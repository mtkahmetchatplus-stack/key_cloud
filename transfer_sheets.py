"""
Transfer T-05-06_SH1 → Sayfa 1 and T-05_06_SH2 → Sayfa 2
Values + formatting copied as-is. Merged cells are unmerged so the
page aspect ratio is preserved without cell-span distortion.
"""

import copy
import openpyxl

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


def transfer_sheet(src_ws, dst_ws):
    # 1. Unmerge all existing merges in destination
    for merge in list(dst_ws.merged_cells.ranges):
        dst_ws.unmerge_cells(str(merge))

    # 2. Clear destination values (now safe — no MergedCell proxies remain)
    for row in dst_ws.iter_rows():
        for cell in row:
            cell.value = None

    # 3. Unmerge source merges too, so we work with plain cells
    src_merges = [str(m) for m in src_ws.merged_cells.ranges]
    for m in src_merges:
        src_ws.unmerge_cells(m)

    # 4. Copy row dimensions
    for row_idx, rd in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_idx].height        = rd.height
        dst_ws.row_dimensions[row_idx].hidden         = rd.hidden
        dst_ws.row_dimensions[row_idx].outline_level  = rd.outline_level

    # 5. Copy column dimensions
    for col_letter, cd in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width         = cd.width
        dst_ws.column_dimensions[col_letter].hidden         = cd.hidden
        dst_ws.column_dimensions[col_letter].outline_level  = cd.outline_level

    # 6. Copy values and styles cell by cell
    for row in src_ws.iter_rows():
        for src_cell in row:
            dst_cell = dst_ws.cell(row=src_cell.row, column=src_cell.column)
            dst_cell.value = src_cell.value
            copy_cell_style(src_cell, dst_cell)

    # 7. Copy page setup
    dst_ws.page_setup = copy.copy(src_ws.page_setup)
    dst_ws.print_area = src_ws.print_area


def main():
    print("Loading workbooks …")
    src_wb = openpyxl.load_workbook(SRC_PATH)
    dst_wb = openpyxl.load_workbook(DST_PATH)

    for src_name, dst_name in SHEET_MAP:
        print(f"  {src_name}  →  {dst_name}")
        src_ws = src_wb[src_name]
        dst_ws = dst_wb[dst_name]
        transfer_sheet(src_ws, dst_ws)
        print(f"    src size : {src_ws.max_row}r × {src_ws.max_column}c")
        print(f"    dst after: {dst_ws.max_row}r × {dst_ws.max_column}c  merged=0 (all unmerged)")

    dst_wb.save(OUT_PATH)
    print(f"\nSaved → {OUT_PATH}")


if __name__ == "__main__":
    main()
