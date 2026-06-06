"""
plan_reader.py — 個別支援計画書（Excel）の内容をテキストに変換
"""
from __future__ import annotations

import io

import openpyxl


def read_plan(excel_bytes: bytes) -> str:
    """
    個別支援計画書のExcelバイト列を読み込み、GPTに渡せるテキストを返す。
    全シート・全セルの内容を結合する。空セル・None は除外する。
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    except Exception as e:
        raise ValueError(
            f"Excelファイルの読み込みに失敗しました: {e}\n"
            "対応形式は .xlsx です。パスワード保護されていないか確認してください。"
        ) from e

    lines: list[str] = []

    for sheet in wb.worksheets:
        sheet_lines: list[str] = []
        for row in sheet.iter_rows():
            row_values: list[str] = []
            for cell in row:
                value = cell.value
                if value is None:
                    continue
                text = str(value).strip()
                if text:
                    row_values.append(text)
            if row_values:
                sheet_lines.append("　".join(row_values))

        if sheet_lines:
            lines.append(f"【{sheet.title}】")
            lines.extend(sheet_lines)

    return "\n".join(lines)
