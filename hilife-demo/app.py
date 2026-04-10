"""
ハイライフ モニタリング報告書 自動生成デモ
面談音声 → 文字起こし → モニタリング報告書整形 → Word/Excel ダウンロード
ブラウザ録音 or ファイルアップロード、新規作成 or 既存ファイルへの追記 に対応
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
    page_title=f"{CLIENT_NAME} | モニタリング報告書 自動生成",
    page_icon="📋",
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
    .stButton > button:hover {{ background-color: #3a5ce4; color: white; }}
    .stButton > button:disabled {{ background-color: #cccccc !important; color: #666 !important; }}
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
    f"<h1 style='color:{PRIMARY_COLOR};text-align:center'>📋 モニタリング報告書 自動生成</h1>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align:center;color:#666'>面談音声をアップロードするだけで、モニタリング報告書をWord/Excelで出力します</p>",
    unsafe_allow_html=True,
)
st.info("⚠️ **1ファイルの上限は25MB（目安：約30分以内）です。**\n\n1時間の面談は前半・後半に分けて、2回アップロードしてください。", icon=None)
st.divider()

# ─── STEP 1: 音声入力（録音 or アップロード） ────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 1</b>　音声を用意する</div>",
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
        pause_threshold=120.0,
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
            st.error(f"ファイルサイズが上限({MAX_FILE_MB}MB)を超えています。")
        else:
            audio_bytes = uploaded.getvalue()
            audio_filename = uploaded.name

# ─── STEP 2: 既存ファイルへの追記（任意） ─────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 2</b>　既存ファイルに追記する場合はアップロード（任意）</div>",
    unsafe_allow_html=True,
)

col_w, col_e = st.columns(2)
with col_w:
    existing_word = st.file_uploader(
        "既存の Word ファイル（追記用）",
        type=["docx"],
        key="existing_word",
    )
with col_e:
    existing_excel = st.file_uploader(
        "既存の Excel ファイル（追記用）",
        type=["xlsx"],
        key="existing_excel",
    )

if existing_word or existing_excel:
    st.success("✅ 既存ファイルを受け取りました。生成後に末尾へ追記します。")

# ─── STEP 3: 実行ボタン ────────────────────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 3</b>　報告書を生成</div>",
    unsafe_allow_html=True,
)

run_btn = st.button("🚀 モニタリング報告書を生成する", use_container_width=True, disabled=audio_bytes is None)

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

    # 報告書整形
    with st.spinner("モニタリング報告書を整形中..."):
        try:
            minutes = summarize(transcript)
        except Exception as e:
            st.error(f"報告書整形に失敗しました: {e}")
            st.stop()

    st.success("✅ 報告書整形完了")

    # ─── プレビュー ────────────────────────────────────────────
    st.divider()
    st.markdown("### 📋 モニタリング報告書 プレビュー")

    col1, col2 = st.columns(2)
    col1.markdown(f"**利用者氏名**: {minutes.get('利用者氏名') or '—'}")
    col2.markdown(f"**作成日**: {minutes.get('作成日') or '—'}")
    col1.markdown(f"**作成者**: {minutes.get('作成者氏名') or '—'}")

    achievement = minutes.get("達成状況の評価", {})
    判定 = achievement.get("判定", "") if isinstance(achievement, dict) else str(achievement)
    badge = {"達成": "🟢", "一部達成": "🟡", "未達成": "🔴"}.get(判定, "⚪")
    st.markdown(f"**達成状況**: {badge} {判定}")

    PREVIEW_FIELDS = [
        ("全体の状況",         minutes.get("全体の状況", "")),
        ("本人の感想・満足度",  minutes.get("本人の感想・満足度", "")),
        ("到達目標",           minutes.get("到達目標", "")),
        ("達成状況（詳細）",   achievement.get("詳細", "") if isinstance(achievement, dict) else ""),
        ("達成されない原因",   minutes.get("達成されない原因の分析", "")),
        ("今後の対応",         minutes.get("今後の対応", "")),
        ("その他留意事項",     minutes.get("その他留意事項", "")),
    ]
    for label, value in PREVIEW_FIELDS:
        if value and value != "該当なし":
            with st.expander(f"▶ {label}", expanded=False):
                st.write(value)

    # ─── STEP 4: ダウンロード ──────────────────────────────────
    st.divider()
    st.markdown(
        "<div class='step-box'><b>STEP 4</b>　ダウンロード</div>",
        unsafe_allow_html=True,
    )

    name = minutes.get("利用者氏名", "").replace(" ", "") or "報告書"
    create_date = (minutes.get("作成日") or "").replace("年", "").replace("月", "").replace("日", "")
    file_stem = f"モニタリング報告書_{name}_{create_date}" if create_date else f"モニタリング報告書_{name}"

    existing_word_bytes = existing_word.getvalue() if existing_word else None
    existing_excel_bytes = existing_excel.getvalue() if existing_excel else None

    col_word, col_excel = st.columns(2)
    with col_word:
        word_bytes = to_word(minutes, existing_bytes=existing_word_bytes)
        label_w = "📄 Word に追記してDL" if existing_word else "📄 Word でダウンロード"
        st.download_button(
            label=label_w,
            data=word_bytes,
            file_name=f"{file_stem}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with col_excel:
        excel_bytes = to_excel(minutes, existing_bytes=existing_excel_bytes)
        label_e = "📊 Excel に追記してDL" if existing_excel else "📊 Excel でダウンロード"
        st.download_button(
            label=label_e,
            data=excel_bytes,
            file_name=f"{file_stem}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ─── フッター ──────────────────────────────────────────────────
st.divider()
st.markdown(
    f"<p style='text-align:center;color:#999;font-size:0.8rem'>{CLIENT_NAME} × AIフローアーキテクト　|　Powered by OpenAI Whisper & GPT-4o</p>",
    unsafe_allow_html=True,
)
