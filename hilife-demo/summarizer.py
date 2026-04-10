"""
summarizer.py — GPT-4o で文字起こしテキストを議事録 JSON に整形
"""
from __future__ import annotations

import json
import os

from openai import OpenAI

SYSTEM_PROMPT = """あなたは介護施設のスタッフと利用者の面談記録を整理するアシスタントです。
与えられた文字起こしテキストを以下のJSON形式に整形してください。

{
  "面談日": "（テキストから読み取れる場合は記載、不明なら空文字）",
  "参加者": ["発言者A", "発言者B"],
  "発言内容": [
    {"話者": "発言者A", "内容": "..."},
    {"話者": "発言者B", "内容": "..."}
  ],
  "要約": "面談全体の要約を200字以内で",
  "確認事項": ["確認・TODO項目を箇条書きで（なければ空リスト）"]
}

- 話者が特定できない場合は「発言者A」「発言者B」のようにラベルを付けてください
- 発言内容はなるべく正確に残してください（要約しすぎない）
- 日本語で出力してください
- JSONのみ返してください（前後に説明文は不要）"""


def summarize(transcript: str) -> dict:
    """
    文字起こしテキストを受け取り、議事録 dict を返す。
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
