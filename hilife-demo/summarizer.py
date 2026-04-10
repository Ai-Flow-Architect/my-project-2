"""
summarizer.py — GPT-4o で文字起こしテキストをモニタリング報告書 JSON に整形
（障害福祉サービス：共同生活援助 / 就労継続支援B型 対応）
"""
from __future__ import annotations

import json
import os

from openai import OpenAI

SYSTEM_PROMPT = """あなたは障害福祉施設（グループホーム・就労継続支援B型）のサービス管理責任者を補佐するアシスタントです。
与えられた面談の文字起こしテキストを、モニタリング報告書として以下のJSON形式に整形してください。

{
  "利用者氏名": "（テキストから読み取れる場合は記載、不明なら空文字）",
  "作成日": "（テキストから読み取れる場合は記載、不明なら空文字）",
  "全体の状況": "生活・就労の現状を200字以内でまとめる",
  "本人の感想・満足度": "利用者自身の発言・気持ち・満足度をそのまま引用・要約する",
  "到達目標": "面談中に言及された目標・希望・支援計画上の目標（複数あれば箇条書き形式で「。」区切り）",
  "達成状況の評価": {
    "判定": "達成 / 一部達成 / 未達成 のいずれか",
    "詳細": "判定の根拠・具体的な状況を記載"
  },
  "達成されない原因の分析": "未達成・一部達成の場合に原因を記載（達成の場合は「該当なし」）",
  "今後の対応": "次の支援方針・変更点・フォロー内容（複数あれば「。」区切り）",
  "作成者氏名": "（テキストからスタッフ名が読み取れる場合は記載、不明なら空文字）",
  "その他留意事項": "特記事項・申し送り・医療的配慮など（なければ「特になし」）"
}

ルール：
- 利用者の発言は「本人の感想・満足度」に優先的に使う
- スタッフの発言は「今後の対応」「全体の状況」に優先的に使う
- 判定は面談内容から総合的に判断する（目標達成の言及がなければ「一部達成」とする）
- JSONのみ返してください（前後に説明文は不要）
- 日本語で出力してください"""


def summarize(transcript: str) -> dict:
    """
    文字起こしテキストを受け取り、モニタリング報告書 dict を返す。
    """
    api_key = os.environ.get("OPENAI_API_KEY", "")
    client = OpenAI(api_key=api_key)

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": transcript},
        ],
        temperature=0.2,
        response_format={"type": "json_object"},
    )

    raw = response.choices[0].message.content or "{}"
    return json.loads(raw)
