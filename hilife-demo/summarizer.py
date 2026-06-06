"""
summarizer.py — GPT-4o で文字起こしテキストをモニタリング報告書 JSON に整形
（障害福祉サービス：共同生活援助 / 就労継続支援B型 対応）
国の基準・実地指導対策に準拠した17項目フォーマット
"""
from __future__ import annotations

import json
import os

from openai import OpenAI

# GPTに渡すテキストの上限（トークン超過防止）
MAX_PLAN_CHARS = 8_000
MAX_TRANSCRIPT_CHARS = 24_000
GPT_TIMEOUT = 120  # 秒

SYSTEM_PROMPT = """あなたは障害福祉施設（グループホーム・就労継続支援B型）のサービス管理責任者を補佐するアシスタントです。
与えられた面談の文字起こしテキストを、モニタリング報告書として以下のJSON形式に整形してください。

{
  "A_基本情報": {
    "利用者氏名": "（テキストから読み取れる場合は記載、不明なら空文字）",
    "モニタリング実施日": "（テキストから読み取れる場合は記載、不明なら空文字）",
    "作成者氏名": "（スタッフ名が読み取れる場合は記載、不明なら空文字）",
    "前回モニタリング実施日": "（言及があれば記載、不明なら空文字）",
    "個別支援計画作成日": "（言及があれば記載、不明なら空文字）"
  },
  "B_支援計画": {
    "支援の全体方針": "面談から読み取れる支援の基本方針・施設の関わり方を記載",
    "長期目標": "1年程度を見据えた目標（面談から読み取れる範囲で）",
    "短期目標": "3〜6ヶ月以内に達成を目指す目標（面談から読み取れる範囲で）"
  },
  "C_モニタリング結果": {
    "全体の状況": "生活・就労の現状を200字以内でまとめる",
    "本人の感想・満足度": "利用者自身の発言・気持ち・満足度をそのまま引用・要約する",
    "家族・保護者の意向": "（言及があれば記載、なければ「面談中言及なし」）",
    "達成状況の評価": {
      "判定": "達成 / 一部達成 / 未達成 のいずれか",
      "詳細": "判定の根拠・具体的な状況を記載"
    },
    "達成されない原因の分析": "未達成・一部達成の場合に原因を記載（達成の場合は「該当なし」）",
    "今後の対応・支援方針": "次の支援方針・変更点・フォロー内容（複数あれば「。」区切り）",
    "計画変更の要否": "要 / 不要 のいずれか（変更が必要な内容が出た場合は「要」）",
    "次回モニタリング予定日": "（言及があれば記載、なければ「実施日から6ヶ月以内」）",
    "その他留意事項": "特記事項・医療的配慮・申し送り（なければ「特になし」）"
  }
}

ルール：
- 利用者の発言は「本人の感想・満足度」に優先的に使う
- スタッフの発言は「今後の対応・支援方針」「全体の状況」に優先的に使う
- 判定は面談内容から総合的に判断する（目標達成の言及がなければ「一部達成」とする）
- 支援の全体方針・長期目標・短期目標は面談内容から読み取れる範囲で推定して記載する
- 計画変更の要否：新しい目標・支援内容・頻度の変更が出た場合は「要」とする
- JSONのみ返してください（前後に説明文は不要）
- 日本語で出力してください"""


def summarize(transcript: str, plan_text: str = "") -> dict:
    """
    文字起こしテキストを受け取り、モニタリング報告書 dict を返す。
    plan_text が渡された場合は個別支援計画書の内容も参照して生成する。
    """
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY が設定されていません。環境変数を確認してください。")
    client = OpenAI(api_key=api_key)

    # トークン超過防止: 長すぎる場合はトランケート
    if len(plan_text) > MAX_PLAN_CHARS:
        plan_text = plan_text[:MAX_PLAN_CHARS] + "\n…（計画書が長いため省略しました）"
    if len(transcript) > MAX_TRANSCRIPT_CHARS:
        transcript = transcript[:MAX_TRANSCRIPT_CHARS] + "\n…（文字起こしが長いため省略しました）"

    if plan_text:
        user_content = (
            f"【個別支援計画書（参考情報）】\n{plan_text}\n\n"
            f"【面談の文字起こし】\n{transcript}"
        )
    else:
        user_content = transcript

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
        temperature=0.2,
        response_format={"type": "json_object"},
        timeout=GPT_TIMEOUT,
    )

    raw = response.choices[0].message.content or "{}"
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"GPTのレスポンスをJSONとして解析できませんでした。\n"
            f"原文（先頭200字）: {raw[:200]}\nエラー: {e}"
        ) from e
