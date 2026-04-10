"""
ハイライフ 音声議事録自動化デモ
面談音声(mp3/m4a/wav) → 文字起こし → 議事録整形 → Word/Excel ダウンロード
ブラウザ録音 or ファイルアップロードの2方式対応
"""
from __future__ import annotations

import os

import streamlit as st
from audio_recorder_streamlit import audio_recorder

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
        return True

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
    .stButton > button:disabled {{
        background-color: #cccccc !important;
        color: #666666 !important;
    }}
    .step-box {{
        background: #f8f9ff;
        border-left: 4px solid {PRIMARY_COLOR};
        border-radius: 4px;
        padding: 0.75rem 1rem;
        margin-bottom: 1rem;
    }}
    .rec-hint {{
        font-size: 0.85rem;
        color: #888;
        margin-top: 0.25rem;
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
st.info("⚠️ **1ファイルの上限は25MB（目安：約30分以内）です。**\n\n1時間の面談は前半・後半に分けて、2回アップロードしてください。", icon=None)
st.divider()

# ─── STEP 1: 音声入力（録音 or アップロード） ────────────────────
st.markdown(
    f"<div class='step-box'><b>STEP 1</b>　音声を用意する</div>",
    unsafe_allow_html=True,
)

tab_record, tab_upload = st.tabs(["🎤 その場で録音する", "📁 ファイルをアップロード"])

audio_bytes: bytes | None = None
audio_filename = "recording.wav"

with tab_record:
    st.markdown("マイクボタンを押して録音を開始 → もう一度押すと停止します")
    recorded = audio_recorder(
        text="",
        recording_color="#e74c3c",
        neutral_color=PRIMARY_COLOR,
        icon_name="microphone",
        icon_size="3x",
        pause_threshold=120.0,  # 2分間無音で自動停止
        sample_rate=16000,
    )
    if recorded:
        audio_bytes = recorded
        audio_filename = "recording.wav"
        st.audio(recorded, format="audio/wav")
        size_mb = len(recorded) / (1024 * 1024)
        st.caption(f"録音サイズ: {size_mb:.1f} MB")
        if size_mb > MAX_FILE_MB:
            st.error(f"録音が{MAX_FILE_MB}MBを超えました。前半・後半に分けてください。")
            audio_bytes = None

with tab_upload:
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
        else:
            audio_bytes = uploaded.getvalue()
            audio_filename = uploaded.name

# ─── STEP 2: 出力形式（両方生成） ────────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 2</b>　出力形式（Word・Excel 両方ダウンロードできます）</div>",
    unsafe_allow_html=True,
)

# ─── STEP 3: 実行ボタン ────────────────────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 2</b>　議事録を生成</div>",
    unsafe_allow_html=True,
)

run_btn = st.button("🚀 議事録を生成する", use_container_width=True, disabled=audio_bytes is None)

if run_btn and audio_bytes:
    # 文字起こし
    with st.spinner("音声を文字起こし中... (1〜2分かかります)"):
        try:
            transcript = transcribe(audio_bytes, audio_filename)
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

    # ─── STEP 3: ダウンロード ──────────────────────────────────
    st.divider()
    st.markdown(
        "<div class='step-box'><b>STEP 3</b>　ダウンロード</div>",
        unsafe_allow_html=True,
    )

    interview_date = minutes.get("面談日", "面談記録").replace("年", "").replace("月", "").replace("日", "")

    col_word, col_excel = st.columns(2)
    with col_word:
        word_bytes = to_word(minutes)
        st.download_button(
            label="📄 Word でダウンロード",
            data=word_bytes,
            file_name=f"面談記録_{interview_date}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with col_excel:
        excel_bytes = to_excel(minutes)
        st.download_button(
            label="📊 Excel でダウンロード",
            data=excel_bytes,
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
