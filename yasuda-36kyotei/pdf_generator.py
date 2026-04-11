"""
pdf_generator.py — 36協定書 HTML→weasyprint PDF生成
サンプルPDF（協定書A社〜F社）レイアウトに準拠
"""
from __future__ import annotations
from pathlib import Path
import weasyprint
from weasyprint.text.fonts import FontConfiguration


# ═══════════════════════════════════════════════════
# CSS（GPT/Gemini推奨 + サンプル実測値）
# ═══════════════════════════════════════════════════
_CSS = """
@page {
    size: A4;
    margin: 14mm 10.5mm 18mm 20mm;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Yu Mincho', '游明朝', 'MS Mincho', 'ＭＳ 明朝',
                 'MS PMincho', 'ＭＳ Ｐ明朝', serif;
    font-size: 9pt;
    line-height: 2.0;
    color: #000;
}
h1 {
    font-size: 14pt;
    font-weight: bold;
    text-align: center;
    line-height: 1.0;
    margin-bottom: 11pt;
    letter-spacing: 1pt;
}
.intro {
    text-align: justify;
    margin-bottom: 0pt;
    font-size: 9pt;
    line-height: 2.0;
}
.article {
    margin-bottom: 0pt;
    font-size: 9pt;
}
.article p {
    text-align: justify;
    line-height: 2.0;
    padding-left: 4em;
    text-indent: -4em;
}
table {
    width: 100%;
    border-collapse: collapse;
    margin: 2pt 0 5pt 0;
    font-size: 9pt;
    table-layout: fixed;
    line-height: 1.3;
}
th, td {
    border: 0.8px solid #000;
    padding: 2pt 3pt;
    vertical-align: middle;
    text-align: center;
    word-break: break-all;
    overflow-wrap: break-word;
    line-height: 1.3;
}
th {
    background-color: #ffffff;
    font-weight: bold;
    font-size: 9pt;
    line-height: 1.3;
}
td.tl { text-align: left; }
.sign-section { margin-top: 10pt; }
.sign-table { width: 100%; border-collapse: collapse; font-size: 9.5pt; }
.sign-table td { border: none; padding: 1pt 2pt; line-height: 1.9; }
.sign-spacer { width: 42%; }
.sign-content { width: 58%; }
"""


# ═══════════════════════════════════════════════════
# ヘルパー
# ═══════════════════════════════════════════════════
def _v(r: dict, key: str, default: str = "") -> str:
    val = r.get(key, default)
    return str(val).strip() if val else default


def _start_date(r: dict) -> str:
    y = _v(r, "起算日_年")
    m = _v(r, "起算日_月")
    return f"令和{y}年{m}月1日"


def _has_special(r: dict) -> bool:
    pat = _v(r, "様式パターン")
    return pat in ("9_2", "9_3", "9_4", "9_5") or "■" in _v(r, "特別条項の有無")


def _is_1nen(r: dict) -> bool:
    return _v(r, "様式パターン") in ("10", "10_2")


# ═══════════════════════════════════════════════════
# テーブル生成（サンプルPDFに完全準拠）
# ═══════════════════════════════════════════════════
def _overtime_table(r: dict) -> str:
    m = _v(r, "起算日_月")
    y = _v(r, "起算日_年")
    月起算 = "毎月1日"
    年起算 = f"毎年{m}月1日"
    期間  = _v(r, "時間外_期間")
    limit = "320" if _is_1nen(r) else "360"

    return f"""
<table>
  <colgroup>
    <col style="width:13%">
    <col style="width:22%">
    <col style="width:13%">
    <col style="width:13%">
    <col style="width:8%">
    <col style="width:12%">
    <col style="width:11%">
    <col style="width:8%">
  </colgroup>
  <thead>
    <tr>
      <th rowspan="3"></th>
      <th rowspan="3">時間外労働をさせる必要のある具体的事由</th>
      <th rowspan="3">業務の種類</th>
      <th rowspan="3">従事する<br>労働者数<br>（満18歳<br>以上の者）</th>
      <th colspan="3">延長することができる時間</th>
      <th rowspan="3">期間</th>
    </tr>
    <tr>
      <th rowspan="2">1日</th>
      <th colspan="2">１日を超える一定期間（起算日）</th>
    </tr>
    <tr>
      <th>１ヶ月<br>毎月１日</th>
      <th>１年<br>毎年{m}月１日</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td style="font-size:8pt; line-height:1.3;">①下記の②に<br>該当しない<br>労働者</td>
      <td class="tl">{_v(r,'時間外_事由')}</td>
      <td class="tl">{_v(r,'時間外_業務の種類')}</td>
      <td>{_v(r,'労働者数')}人</td>
      <td>{_v(r,'延長時間_1日')}時間</td>
      <td>{_v(r,'延長時間_1ヶ月')}時間</td>
      <td>{limit}時間</td>
      <td class="tl">{期間}</td>
    </tr>
    <tr style="height:22pt;">
      <td style="font-size:8pt; line-height:1.3;">②1年単位の<br>変形労働時間制<br>により労働する<br>労働者</td>
      <td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    </tr>
  </tbody>
</table>"""


def _holiday_table(r: dict) -> str:
    return f"""
<table>
  <colgroup>
    <col style="width:26%">
    <col style="width:13%">
    <col style="width:9%">
    <col style="width:12%">
    <col style="width:20%">
    <col style="width:20%">
  </colgroup>
  <thead>
    <tr>
      <th>休日労働をさせる<br>必要のある具体的事由</th>
      <th>業務の種類</th>
      <th>労働者数<br>（満18歳<br>以上の者）</th>
      <th>所定<br>休日</th>
      <th>労働させることができる休日<br>並びに始業及び終業の時刻</th>
      <th>期間</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="tl">{_v(r,'休日_事由')}</td>
      <td class="tl">{_v(r,'休日_業務の種類')}</td>
      <td>{_v(r,'労働者数')}人</td>
      <td>{_v(r,'所定休日')}</td>
      <td>1か月に{_v(r,'休日労働_日数')}日<br>{_v(r,'始業終業時刻')}</td>
      <td class="tl">{_v(r,'休日_期間')}</td>
    </tr>
  </tbody>
</table>"""


def _special_table(r: dict) -> str:
    if not _has_special(r):
        return ""
    措置 = _v(r, "特別_健康措置_内容") or f"番号{_v(r,'特別_健康措置_番号')}の措置"
    return f"""
<div class="article" style="margin-top:4pt;">
<p>（特別条項）特別の事情がある場合、次のとおり時間外労働を延長させることができる。</p>
<table>
  <colgroup>
    <col style="width:22%"><col style="width:13%"><col style="width:9%">
    <col style="width:9%"><col style="width:8%">
    <col style="width:10%"><col style="width:9%"><col style="width:10%">
  </colgroup>
  <thead>
    <tr>
      <th>臨時的に限度時間を超えて<br>労働させる必要のある<br>具体的事由</th>
      <th>業務の種類</th>
      <th>労働者数<br>（満18歳<br>以上）</th>
      <th>１回につき<br>延長できる<br>時間</th>
      <th>延長<br>できる<br>回数</th>
      <th>延長時間<br>（月）</th>
      <th>割増<br>賃金率</th>
      <th>延長時間<br>（年）</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td class="tl">{_v(r,'特別_理由')}</td>
      <td class="tl">{_v(r,'特別_業務の種類')}</td>
      <td>{_v(r,'特別_労働者数')}人</td>
      <td>{_v(r,'特別_延長時間')}時間</td>
      <td>{_v(r,'特別_超過回数')}回</td>
      <td>{_v(r,'特別_延長時間_月')}時間</td>
      <td>{_v(r,'特別_割増賃金率')}%</td>
      <td>{_v(r,'特別_延長時間_年')}時間</td>
    </tr>
  </tbody>
</table>
<p>上記で定める時間を超えて労働させる場合の手続き：{_v(r,'特別_手続き')}</p>
<p>上限時間を超えた場合の健康確保措置：{措置}</p>
</div>"""


# ═══════════════════════════════════════════════════
# HTML 組み立て
# ═══════════════════════════════════════════════════
def _build_html(r: dict) -> str:
    社名    = _v(r, "事業所名")
    職名    = _v(r, "事業主職名", "代表取締役")
    代表者  = _v(r, "事業主名")
    代表職  = _v(r, "労働者代表_職")
    代表氏名 = _v(r, "労働者代表_氏名")
    起算日  = _start_date(r)
    年号年  = _v(r, "起算日_年")

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>{_CSS}</style>
</head>
<body>

<h1>時間外労働及び休日労働に関する協定書</h1>

<p class="intro">{社名}（以下「甲」という。）と労働者代表者（以下「乙」という。）は、労働基準法第３６条第１項の規定に基づき、労働基準法に定める法定労働時間（１週４０時間、１日８時間）並びに変形労働時間制に定める所定労働時間を超えた労働時間で、かつ１日８時間、１週４０時間の法定労働時間又は変形期間の法定労働時間の総枠を超える労働（以下「時間外労働」という。）及び労働基準法に定める休日（毎週１日又は４週４日）における労働（以下「休日労働」という。）に関し、次の通り協定する。</p>

<div class="article">
<p><strong>第１条</strong>　甲は、時間外労働及び休日労働を可能な限り行わせないように努める。</p>
</div>

<div class="article">
<p><strong>第２条</strong>　乙は、故意または過失により時間外労働及び休日労働を生じさせない義務を負う。</p>
</div>

<div class="article">
<p><strong>第３条</strong>　前2条にも関わらずその必要性を生じた場合、甲は次により時間外労働を行わせることができる。</p>
{_overtime_table(r)}
</div>

{_special_table(r)}

<div class="article">
<p><strong>第４条</strong>　甲は、時間外労働を行わせる場合は、原則として、前日の終業時刻までに当該労働者に通知する。また、休日労働を行わせる場合は、原則として、前日の終業時刻までに当該労働者に通知する。</p>
</div>

<div class="article">
<p><strong>第５条</strong>　第２条の表における1週、１ヶ月及び１年の起算日はいずれも{起算日}とする。</p>
</div>

<div class="article">
<p><strong>第６条</strong>　甲は、必要がある場合には、次により休日労働を行わせることができる。</p>
{_holiday_table(r)}
</div>

<div class="article">
<p><strong>第７条</strong>　本協定の有効期間は、{起算日}から１年間とする。</p>
</div>

<div class="sign-section">
  <div style="margin-left:42%; font-size:10.5pt; line-height:1.9;">
    <div style="text-align:right; padding-right:2pt; margin-bottom:6pt;">令和　{年号年}　年　　　月　　　日</div>
    <div style="margin-bottom:0pt;">（甲）{社名}</div>
    <div style="text-align:right; padding-right:2pt; margin-bottom:8pt;">{職名}　{代表者}</div>
    <table style="width:100%; border-collapse:separate; border-spacing:0; font-size:10.5pt; line-height:2.0;">
      <colgroup>
        <col style="width:110pt;">
        <col>
      </colgroup>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:left; padding:0 4pt 0 0;">（乙）労働者代表　職種</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表職}</td>
      </tr>
      <tr>
        <td style="border:none; white-space:nowrap; vertical-align:bottom; text-align:right; padding:3pt 4pt 0 0;">署名</td>
        <td style="border:none; border-bottom:1px solid #000; vertical-align:bottom; text-align:left; padding:1pt 6pt;">{代表氏名}</td>
      </tr>
    </table>
  </div>
</div>

</body>
</html>"""


# ═══════════════════════════════════════════════════
# 公開 API
# ═══════════════════════════════════════════════════
def generate_pdf(record: dict) -> bytes:
    """1件のレコードからPDFバイト列を生成して返す"""
    html = _build_html(record)
    fc = FontConfiguration()
    return weasyprint.HTML(string=html).write_pdf(font_config=fc)


def generate_pdf_file(record: dict, output_dir: str = "output") -> str:
    """PDFをファイルに保存してパスを返す"""
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    社名 = _v(record, "事業所名").replace(" ", "").replace("　", "")
    pat  = _v(record, "様式パターン", "9")
    out  = Path(output_dir) / f"36協定書_{社名}_{pat}.pdf"
    out.write_bytes(generate_pdf(record))
    return str(out)
