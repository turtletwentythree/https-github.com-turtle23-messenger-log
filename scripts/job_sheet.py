#!/usr/bin/env python3
"""
Build the daily Messenger job-sheet workbook (4 sheets) from already-parsed
booking records — a direct port of the final, user-approved
generate_job_sheet_v3.py logic, adapted to take record dicts (as produced by
sync_data.load_records) instead of raw CSV rows.

Sheets:
  ส่งเอกสารช่วงเช้า / ส่งเอกสารช่วงบ่าย
      Auto-filled from the queue — ALL bookings for that date+shift.
  รับฝากเอกสารกลับเช้า / รับฝากเอกสารกลับบ่าย
      Blank hand-log template (date auto-updates), padded + scaled to fill
      one printed page.
"""
from io import BytesIO

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

SEND_HEADERS = [
    "No.", "วันที่", "แผนก", "ชื่อผู้สั่งงาน", "เบอร์โทรติดต่อ", "ประเภทการสั่งงาน",
    "สถานที่", "ชื่อผู้ติดต่อปลายทาง", "เบอร์โทรผู้ติดต่อปลายทาง",
    "รายละเอียดเพิ่มเติม (ถ้ามี)", "ต้องการส่งเอกสารให้ปลายทางในช่วงใด", "เซ็นรับเอกสาร",
]
SEND_WIDTHS = [5, 11, 13, 18, 14, 16, 22, 20, 16, 26, 13, 14]

RETURN_HEADERS = [
    "No.", "วันที่", "ชื่อผู้ฝากเอกสารกลับ", "เบอร์โทรติดต่อ", "ฝากเอกสารจากสถานที่",
    "รายละเอียดเอกสารเบื้องต้น", "เอกสารฝากถึงใคร", "จำนวน",
    "ชื่อผู้รับฝากเอกสารกลับ", "ชื่อผู้รับเอกสารจริง",
]
RETURN_WIDTHS = [5, 11, 18, 14, 20, 22, 18, 8, 18, 18]
RETURN_BLANK_ROWS = 20  # padded + fitToHeight=1 below so the sheet fills one full printed page

SEND_BLANK_ROWS = 20  # when a date/shift has 0 queue bookings, still print this many
# blank numbered rows below the "no data" note so the sheet can be filled in by hand

SHIFTS = ["เช้า", "บ่าย"]
SEND_SHEET_NAME = {"เช้า": "ส่งเอกสารช่วงเช้า", "บ่าย": "ส่งเอกสารช่วงบ่าย"}
RETURN_SHEET_NAME = {"เช้า": "รับฝากเอกสารกลับเช้า", "บ่าย": "รับฝากเอกสารกลับบ่าย"}
SEND_TIME = {"เช้า": "10.30 น.", "บ่าย": "14.30 น."}
RETURN_TIME = {"เช้า": "10.00 น.", "บ่าย": "14.00 น."}

THIN = Side(style="thin", color="FF000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E2DC")


def new_sheet(wb, name, ncols):
    ws = wb.create_sheet(name)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:2"
    ws.freeze_panes = "A3"
    return ws


def write_title(ws, title, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = Font(name="Calibri", size=16, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28


def write_header(ws, headers):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.font = Font(name="Calibri", size=10.5, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER
    ws.row_dimensions[2].height = 30


def build_send_sheet(wb, shift, records, target_date, date_str):
    ws = new_sheet(wb, SEND_SHEET_NAME[shift], len(SEND_HEADERS))
    write_title(ws, f"งาน Messenger ช่วง{shift} {SEND_TIME[shift]} วันที่ {date_str}", len(SEND_HEADERS))
    write_header(ws, SEND_HEADERS)

    row = 3
    for i, rec in enumerate(records):
        values = [
            i + 1, target_date, rec["dept"], rec["requester"], rec["phone"],
            rec["task"], rec["loc"], rec["contact"], rec["contact_phone"],
            rec["notes"], "ช่วง" + rec["shift"], None,
        ]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=row, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=(c in (7, 10)))
            if c == 2:
                cell.number_format = "d/m/yyyy"
        ws.row_dimensions[row].height = 24
        row += 1

    if not records:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(SEND_HEADERS))
        cell = ws.cell(
            row=row,
            column=1,
            value="ไม่มีรายการจากระบบจองคิวสำหรับวันและช่วงเวลานี้ — ตารางด้านล่างเว้นไว้ให้เขียนเพิ่มด้วยมือ",
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(italic=True, color="FF808080")
        for c in range(1, len(SEND_HEADERS) + 1):
            ws.cell(row=row, column=c).border = BORDER
        ws.row_dimensions[row].height = 22
        row += 1

    # Blank hand-log rows, same style as the return sheet, appended after
    # whatever's above (real bookings and/or the "no data" note) so there's
    # always room to add jobs that aren't in the queue yet — numbering picks
    # up where the real rows left off rather than restarting at 1.
    start_no = len(records) + 1
    for i in range(SEND_BLANK_ROWS):
        rr = row + i
        ws.cell(row=rr, column=1, value=start_no + i).border = BORDER
        ws.cell(row=rr, column=1).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(2, len(SEND_HEADERS) + 1):
            ws.cell(row=rr, column=c).border = BORDER
        ws.row_dimensions[rr].height = 24

    if not records:
        ws.page_setup.fitToHeight = 1  # scale the (all-blank) hand-log to fill one printed page

    for c, w in enumerate(SEND_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    return ws


def build_return_sheet(wb, shift, date_str):
    ws = new_sheet(wb, RETURN_SHEET_NAME[shift], len(RETURN_HEADERS))
    ws.page_setup.fitToHeight = 1  # blank hand-log — scale to fill exactly one printed page
    write_title(ws, f"รับฝากเอกสารกลับ Messenger ช่วง{shift} {RETURN_TIME[shift]} วันที่ {date_str}", len(RETURN_HEADERS))
    write_header(ws, RETURN_HEADERS)

    for i in range(RETURN_BLANK_ROWS):
        r = 3 + i
        ws.cell(row=r, column=1, value=i + 1).border = BORDER
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(2, len(RETURN_HEADERS) + 1):
            ws.cell(row=r, column=c).border = BORDER
        ws.row_dimensions[r].height = 24

    for c, w in enumerate(RETURN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    return ws


def build_workbook(records, target_date):
    """target_date: a datetime.date. Returns (xlsx_bytes, filename)."""
    date_str = f"{target_date.day}/{target_date.month}/{target_date.year}"
    iso = target_date.isoformat()

    buckets = {s: [] for s in SHIFTS}
    for rec in records:
        if rec["date"] != iso:
            continue
        if rec["shift"] in buckets:
            buckets[rec["shift"]].append(rec)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for shift in SHIFTS:
        build_send_sheet(wb, shift, buckets[shift], target_date, date_str)
    for shift in SHIFTS:
        build_return_sheet(wb, shift, date_str)

    buf = BytesIO()
    wb.save(buf)

    filename = f"MESSENGER {target_date.day}-{target_date.month}-{target_date.year}.xlsx"
    return buf.getvalue(), filename
