"""
ハイライフ 音声議事録自動化デモ
面談音声(mp3/m4a/wav) → 文字起こし → 議事録整形 → Word/Excel ダウンロード
"""
from __future__ import annotations

import os

import streamlit as st

from transcriber import transcribe
from summarizer import summarize
from exporter import to_word, to_excel

# ─── 定数 ──────────────────────────────────────────────────────
CLIENT_NAME = "ハイライフ"
PRIMARY_COLOR = "#4a6cf7"
MAX_FILE_MB = 25
SUPPORTED_TYPES = ["mp3", "m4a", "wav", "webm"]


# ─── パスワード認証 ─────────────────────────────────────────────
def check_password() -> bool:
    correct_pw: str = ""
    try:
        correct_pw = st.secrets["password"]
    except Exception:
        correct_pw = os.environ.get("APP_PASSWORD", "")

    if not correct_pw:
        return True  # 未設定時は開放（ローカル開発用）

    if st.session_state.get("authenticated"):
        return True

    with st.container():
        st.markdown(
            f"<h2 style='text-align:center;color:{PRIMARY_COLOR}'>🔒 パスワードを入力してください</h2>",
            unsafe_allow_html=True,
        )
        pw = st.text_input("パスワード", type="password", key="pw_input")
        if st.button("ログイン", use_container_width=True):
            if pw == correct_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("パスワードが違います")
    return False


# ─── ページ設定 ────────────────────────────────────────────────
st.set_page_config(
    page_title=f"{CLIENT_NAME} | 音声議事録自動化",
    page_icon="🎙️",
    layout="centered",
)

st.markdown(
    f"""
    <style>
    .stButton > button {{
        background-color: {PRIMARY_COLOR};
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.5rem;
        font-weight: bold;
    }}
    .stButton > button:hover {{
        background-color: #3a5ce4;
        color: white;
    }}
    .step-box {{
        background: #f8f9ff;
        border-left: 4px solid {PRIMARY_COLOR};
        border-radius: 4px;
        padding: 0.75rem 1rem;
        margin-bottom: 1rem;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

if not check_password():
    st.stop()

# ─── ヘッダー ──────────────────────────────────────────────────
st.markdown(
    f"<h1 style='color:{PRIMARY_COLOR};text-align:center'>🎙️ 音声議事録自動化</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align:center;color:#666'>面談音声をアップロードするだけで、議事録をWord/Excelで出力します</p>",
    unsafe_allow_html=True,
)
st.divider()

# ─── STEP 1: ファイルアップロード ──────────────────────────────
st.markdown(
    f"<div class='step-box'><b>STEP 1</b>　音声ファイルをアップロード</div>",
    unsafe_allow_html=True,
)
uploaded = st.file_uploader(
    f"対応形式: mp3 / m4a / wav / webm　（上限 {MAX_FILE_MB}MB）",
    type=SUPPORTED_TYPES,
    label_visibility="collapsed",
)

if uploaded:
    file_mb = len(uploaded.getvalue()) / (1024 * 1024)
    st.caption(f"ファイル名: {uploaded.name}　｜　サイズ: {file_mb:.1f} MB")

    if file_mb > MAX_FILE_MB:
        st.error(f"ファイルサイズが上限({MAX_FILE_MB}MB)を超えています。小さいファイルに分割してください。")
        st.stop()

# ─── STEP 2: 出力形式選択 ─────────────────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 2</b>　出力形式を選択</div>",
    unsafe_allow_html=True,
)
output_format = st.radio(
    "出力形式",
    options=["Word (.docx)", "Excel (.xlsx)"],
    horizontal=True,
    label_visibility="collapsed",
)

# ─── STEP 3: 実行ボタン ────────────────────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 3</b>　議事録を生成</div>",
    unsafe_allow_html=True,
)

run_btn = st.button("🚀 議事録を生成する", use_container_width=True, disabled=uploaded is None)

if run_btn and uploaded:
    audio_bytes = uploaded.getvalue()

    # 文字起こし
    with st.spinner("音声を文字起こし中... (1〜2分かかります)"):
        try:
            transcript = transcribe(audio_bytes, uploaded.name)
        except Exception as e:
            st.error(f"文字起こしに失敗しました: {e}")
            st.stop()

    st.success("✅ 文字起こし完了")

    with st.expander("文字起こし全文を確認する", expanded=False):
        st.text_area("テキスト", transcript, height=200, label_visibility="collapsed")

    # 議事録整形
    with st.spinner("議事録を整形中..."):
        try:
            minutes = summarize(transcript)
        except Exception as e:
            st.error(f"議事録整形に失敗しました: {e}")
            st.stop()

    st.success("✅ 議事録整形完了")

    # ─── プレビュー ────────────────────────────────────────────
    st.divider()
    st.markdown("### 📋 議事録プレビュー")

    col1, col2 = st.columns(2)
    col1.markdown(f"**面談日**: {minutes.get('面談日', '—')}")
    col2.markdown(f"**参加者**: {', '.join(minutes.get('参加者', []))}")

    st.markdown("**要約**")
    st.info(minutes.get("要約", ""))

    if minutes.get("確認事項"):
        st.markdown("**確認事項**")
        for item in minutes["確認事項"]:
            st.markdown(f"- {item}")

    with st.expander("発言内容（全件）", expanded=False):
        for item in minutes.get("発言内容", []):
            st.markdown(f"**{item.get('話者', '')}**: {item.get('内容', '')}")

    # ─── STEP 4: ダウンロード ──────────────────────────────────
    st.divider()
    st.markdown(
        "<div class='step-box'><b>STEP 4</b>　ダウンロード</div>",
        unsafe_allow_html=True,
    )

    interview_date = minutes.get("面談日", "面談記録").replace("年", "").replace("月", "").replace("日", "")

    if output_format == "Word (.docx)":
        file_bytes = to_word(minutes)
        st.download_button(
            label="📄 Word ファイルをダウンロード",
            data=file_bytes,
            file_name=f"面談記録_{interview_date}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        file_bytes = to_excel(minutes)
        st.download_button(
            label="📊 Excel ファイルをダウンロード",
            data=file_bytes,
            file_name=f"面談記録_{interview_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ─── フッター ──────────────────────────────────────────────────
st.divider()
st.markdown(
    f"<p style='text-align:center;color:#999;font-size:0.8rem'>{CLIENT_NAME} × AIフローアーキテクト　|　Powered by OpenAI Whisper & GPT-4o</p>",
    unsafe_allow_html=True,
)
