"""
exporter.py — モニタリング報告書 dict を Word(.docx) または Excel(.xlsx) に書き出す
新規作成・既存ファイルへの追記 の両方に対応
"""
from __future__ import annotations

import io
from datetime import date

from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

_HEADER_COLOR = "4A6CF7"
_LIGHT_BG = "EEF2FF"
_WHITE = "FFFFFF"


# ══════════════════════════════════════════════════════════
#  Word 出力
# ══════════════════════════════════════════════════════════

def _add_page_break(doc: Document) -> None:
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    doc.element.body.append(p)


def _write_monitoring_to_doc(doc: Document, minutes: dict) -> None:
    """モニタリング報告書1件分を doc に書き込む（新規・追記共通）"""

    today = date.today().strftime("%Y年%m月%d日")
    create_date = minutes.get("作成日") or today

    # タイトル
    title = doc.add_heading("モニタリング報告書", level=1)
    title.runs[0].font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

    # 基本情報
    doc.add_paragraph(f"利用者氏名：{minutes.get('利用者氏名') or '　　　　'}")
    doc.add_paragraph(f"作成日：{create_date}　　作成者：{minutes.get('作成者氏名') or '　　　　'}")
    doc.add_paragraph("")

    FIELDS = [
        ("全体の状況",         minutes.get("全体の状況", "")),
        ("本人の感想・満足度",  minutes.get("本人の感想・満足度", "")),
        ("到達目標",           minutes.get("到達目標", "")),
        ("達成状況の評価",     _fmt_achievement(minutes.get("達成状況の評価", {}))),
        ("達成されない原因の分析", minutes.get("達成されない原因の分析", "")),
        ("今後の対応",         minutes.get("今後の対応", "")),
        ("その他留意事項",     minutes.get("その他留意事項", "特になし")),
    ]

    for label, value in FIELDS:
        h = doc.add_paragraph()
        run_h = h.add_run(f"■ {label}")
        run_h.bold = True
        run_h.font.size = Pt(10.5)
        run_h.font.color.rgb = RGBColor(0x1a, 0x1a, 0x2e)

        body = doc.add_paragraph(value or "—")
        body.runs[0].font.size = Pt(10.5)
        doc.add_paragraph("")


def _fmt_achievement(v) -> str:
    if isinstance(v, dict):
        判定 = v.get("判定", "")
        詳細 = v.get("詳細", "")
        return f"【{判定}】{詳細}" if 判定 else 詳細
    return str(v) if v else ""


def to_word(minutes: dict, existing_bytes: bytes | None = None) -> bytes:
    """
    existing_bytes が None → 新規作成
    existing_bytes が指定 → 既存 docx の末尾に改ページ＋追記
    """
    if existing_bytes:
        doc = Document(io.BytesIO(existing_bytes))
        _add_page_break(doc)
    else:
        doc = Document()

    _write_monitoring_to_doc(doc, minutes)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════
#  Excel 出力
# ══════════════════════════════════════════════════════════

def _thin_border() -> Border:
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)


def _header_style(cell, text: str) -> None:
    cell.value = text
    cell.font = Font(bold=True, color=_WHITE, size=11)
    cell.fill = PatternFill("solid", fgColor=_HEADER_COLOR)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def _data_style(cell, text: str, bg: bool = False) -> None:
    cell.value = text
    cell.font = Font(size=10)
    cell.fill = PatternFill("solid", fgColor=_LIGHT_BG if bg else _WHITE)
    cell.border = _thin_border()
    cell.alignment = Alignment(vertical="top", wrap_text=True)


def _append_monitoring_to_sheet(ws, minutes: dict, start_row: int) -> int:
    """1件分をシートの start_row から書き込み、次の開始行を返す"""

    today = date.today().strftime("%Y年%m月%d日")
    create_date = minutes.get("作成日") or today

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 65

    row = start_row

    def header_row(label: str) -> None:
        nonlocal row
        c = ws.cell(row=row, column=1, value=label)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        _header_style(c, label)
        ws.row_dimensions[row].height = 20
        row += 1

    def data_row(label: str, value: str, bg: bool = False, height: int = 15) -> None:
        nonlocal row
        c1 = ws.cell(row=row, column=1, value=label)
        c2 = ws.cell(row=row, column=2, value=value)
        c1.font = Font(bold=True, size=10)
        c1.border = _thin_border()
        c1.alignment = Alignment(vertical="top", wrap_text=True)
        c1.fill = PatternFill("solid", fgColor=_LIGHT_BG if bg else _WHITE)
        _data_style(c2, value, bg)
        ws.row_dimensions[row].height = height
        row += 1

    # 基本情報
    header_row("■ 基本情報")
    data_row("利用者氏名",  minutes.get("利用者氏名") or "")
    data_row("作成日",      create_date, bg=True)
    data_row("作成者氏名",  minutes.get("作成者氏名") or "")
    row += 1  # 空行

    # 各項目
    FIELDS = [
        ("全体の状況",            minutes.get("全体の状況", ""),          60),
        ("本人の感想・満足度",     minutes.get("本人の感想・満足度", ""),   60),
        ("到達目標",              minutes.get("到達目標", ""),             40),
        ("達成状況の評価",        _fmt_achievement(minutes.get("達成状況の評価", {})), 40),
        ("達成されない原因の分析", minutes.get("達成されない原因の分析", ""), 40),
        ("今後の対応",            minutes.get("今後の対応", ""),           50),
        ("その他留意事項",        minutes.get("その他留意事項", "特になし"), 30),
    ]

    for i, (label, value, h) in enumerate(FIELDS):
        header_row(f"■ {label}")
        data_row("", value, bg=False, height=h)
        row += 1  # 空行

    return row  # 次回の開始行


def to_excel(minutes: dict, existing_bytes: bytes | None = None) -> bytes:
    """
    existing_bytes が None → 新規作成
    existing_bytes が指定 → 既存 xlsx の末尾に追記
    """
    if existing_bytes:
        wb = openpyxl.load_workbook(io.BytesIO(existing_bytes))
        ws = wb.active
        # 既存の最終行を探して空行2行あけて追記
        max_row = ws.max_row
        start_row = max_row + 3
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "モニタリング報告書"
        start_row = 1

    _append_monitoring_to_sheet(ws, minutes, start_row)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
