"""
exporter.py — 議事録 dict を Word(.docx) または Excel(.xlsx) に書き出す
"""
from __future__ import annotations

import io
from datetime import date

from docx import Document
from docx.shared import Pt, RGBColor
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ─── Word 出力 ───────────────────────────────────────────────

def to_word(minutes: dict) -> bytes:
    doc = Document()

    # タイトル
    title = doc.add_heading("面談記録", level=1)
    title.runs[0].font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

    # 基本情報
    doc.add_paragraph(f"面談日: {minutes.get('面談日') or date.today().strftime('%Y年%m月%d日')}")
    participants = "、".join(minutes.get("参加者", []))
    doc.add_paragraph(f"参加者: {participants}")
    doc.add_paragraph("")

    # 発言内容
    doc.add_heading("■ 発言内容", level=2)
    for item in minutes.get("発言内容", []):
        p = doc.add_paragraph()
        run_speaker = p.add_run(f"【{item.get('話者', '')}】 ")
        run_speaker.bold = True
        run_speaker.font.size = Pt(10.5)
        run_content = p.add_run(item.get("内容", ""))
        run_content.font.size = Pt(10.5)

    doc.add_paragraph("")

    # 要約
    doc.add_heading("■ 要約", level=2)
    doc.add_paragraph(minutes.get("要約", ""))

    # 確認事項
    items = minutes.get("確認事項", [])
    if items:
        doc.add_paragraph("")
        doc.add_heading("■ 確認事項", level=2)
        for item in items:
            doc.add_paragraph(f"・{item}")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─── Excel 出力 ──────────────────────────────────────────────

_HEADER_COLOR = "4A6CF7"
_LIGHT_BG = "F0F4FF"

def _thin_border() -> Border:
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)


def to_excel(minutes: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "面談記録"

    # 列幅
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 70

    row = 1

    def _header_row(label: str) -> None:
        nonlocal row
        cell = ws.cell(row=row, column=1, value=label)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor=_HEADER_COLOR)
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row].height = 20
        row += 1

    def _data_row(label: str, value: str, bg: bool = False) -> None:
        nonlocal row
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=value)
        for c in (c1, c2):
            c.border = _thin_border()
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if bg:
                c.fill = PatternFill("solid", fgColor=_LIGHT_BG)
        c1.font = Font(bold=True, size=10)
        c2.font = Font(size=10)
        ws.row_dimensions[row].height = 15
        row += 1

    # 基本情報
    _header_row("■ 基本情報")
    interview_date = minutes.get("面談日") or date.today().strftime("%Y年%m月%d日")
    _data_row("面談日", interview_date)
    participants = "、".join(minutes.get("参加者", []))
    _data_row("参加者", participants, bg=True)
    row += 1

    # 発言内容
    _header_row("■ 発言内容")
    for i, item in enumerate(minutes.get("発言内容", [])):
        _data_row(item.get("話者", ""), item.get("内容", ""), bg=(i % 2 == 1))
        ws.row_dimensions[row - 1].height = max(15, len(item.get("内容", "")) // 4)
    row += 1

    # 要約
    _header_row("■ 要約")
    summary = minutes.get("要約", "")
    c1 = ws.cell(row=row, column=1, value="要約")
    c2 = ws.cell(row=row, column=2, value=summary)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=1)
    for c in (c1, c2):
        c.border = _thin_border()
        c.alignment = Alignment(vertical="top", wrap_text=True)
    c1.font = Font(bold=True, size=10)
    c2.font = Font(size=10)
    ws.row_dimensions[row].height = 60
    row += 2

    # 確認事項
    items = minutes.get("確認事項", [])
    if items:
        _header_row("■ 確認事項")
        for i, item in enumerate(items):
            _data_row(f"確認{i+1}", item, bg=(i % 2 == 1))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
