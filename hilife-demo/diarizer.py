"""
diarizer.py — 文字起こしテキストを GPT で「話者分解」する（簡易・無料・クラウド完結）
音声からの声紋分離ではなく、会話の文脈から発言者（支援員/利用者/家族 等）を推定してラベル付けする。
追加ベンダー・音声の外部送信なし。OPENAI_API_KEY のみ使用。
"""
from __future__ import annotations

import json
import os

from openai import OpenAI

GPT_TIMEOUT = 120  # 秒
MAX_TRANSCRIPT_CHARS = 24_000

SYSTEM_PROMPT = """あなたは障害福祉施設の面談記録を整理するアシスタントです。
与えられた面談の文字起こしテキストを、発言者ごとに分解してください。
音声の声紋ではなく「会話の流れ・敬語・立場・話題」から発言者を推定します。

出力は次のJSON形式のみ：
{
  "話者ラベル": ["支援員", "利用者", ...],
  "発言": [
    {"話者": "支援員", "発言": "本日もよろしくお願いします。"},
    {"話者": "利用者", "発言": "少し足が痛くて歩くのが大変でした。"}
  ]
}

ルール：
- 話者は実名が判れば実名、不明なら「支援員」「利用者」「家族」「サービス管理責任者」等の役割名にする
- 1つの発言の途中で話者が変わる場合は分割する
- 発言内容は文字起こしのまま（要約・改変しない・脱字補正は最小限）
- 誰の発言か判断がつかない箇所は「話者」を「不明」とする
- JSONのみ返す（前後の説明文は不要）"""


def diarize(transcript: str) -> dict:
    """
    文字起こしテキストを受け取り、話者分解結果 dict を返す。
    返り値: {"話者ラベル": [...], "発言": [{"話者":..,"発言":..}, ...]}
    失敗時も例外を投げず、全文を「不明」1発言として返す（後段を止めない）。
    """
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY が設定されていません。環境変数を確認してください。")

    text = transcript
    if len(text) > MAX_TRANSCRIPT_CHARS:
        text = text[:MAX_TRANSCRIPT_CHARS] + "\n…（文字起こしが長いため省略しました）"

    client = OpenAI(api_key=api_key)
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": text},
            ],
            temperature=0.1,
            response_format={"type": "json_object"},
            timeout=GPT_TIMEOUT,
        )
        raw = response.choices[0].message.content or "{}"
        data = json.loads(raw)
        if not isinstance(data.get("発言"), list) or not data["発言"]:
            raise ValueError("発言リストが空")
        data.setdefault("話者ラベル", sorted({u.get("話者", "不明") for u in data["発言"]}))
        return data
    except Exception:
        # 失敗しても後段（要約）を止めない安全フォールバック
        return {"話者ラベル": ["不明"], "発言": [{"話者": "不明", "発言": transcript}]}


def to_text(diarized: dict) -> str:
    """話者分解 dict を「話者: 発言」形式の表示用テキストに整形する。"""
    lines = []
    for u in diarized.get("発言", []):
        lines.append(f"{u.get('話者', '不明')}：{u.get('発言', '')}")
    return "\n".join(lines)
