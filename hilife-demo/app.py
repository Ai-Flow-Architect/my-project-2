"""
ハイライフ モニタリング報告書 自動生成デモ
面談音声 → 文字起こし → モニタリング報告書整形 → Word/Excel ダウンロード
ブラウザ録音 or ファイルアップロード、新規作成 or 既存ファイルへの追記 に対応
"""
from __future__ import annotations

import hashlib
import os

import streamlit as st
from audio_recorder_streamlit import audio_recorder

from transcriber import transcribe
from summarizer import summarize
from exporter import to_word, to_excel
from plan_reader import read_plan


# ─── Streamlit Secrets → 環境変数ブリッジ ──────────────────────
# transcriber/summarizer は os.environ から OPENAI_API_KEY を読むため、
# Streamlit Cloud の Secrets（st.secrets）を環境変数へ橋渡しする。
def _bridge_secret(key: str) -> None:
    try:
        if key in st.secrets and not os.environ.get(key):
            os.environ[key] = str(st.secrets[key])
    except Exception:
        pass


_bridge_secret("OPENAI_API_KEY")

# ─── 定数 ──────────────────────────────────────────────────────
CLIENT_NAME = "ハイライフ"
PRIMARY_COLOR = "#4a6cf7"
MAX_FILE_MB = 200
SUPPORTED_TYPES = ["mp3", "m4a", "wav", "webm", "mp4"]


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
st.info("📁 **最大2時間（目安200MB）の音声ファイルに対応しています。**\n\n25MBを超える場合は自動的に分割して処理します。mp3 / m4a / wav / webm / mp4 に対応。", icon=None)
st.divider()

# ─── STEP 1: 音声入力（録音 or アップロード） ────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 1</b>　音声を用意する</div>",
    unsafe_allow_html=True,
)

tab_record, tab_upload = st.tabs(["🎤 その場で録音する", "📁 ファイルをアップロード"])

audio_bytes: bytes | None = None
audio_filename = "recording.wav"
audio_source = ""  # "record" or "upload"

with tab_record:
    st.markdown("**🎤 マイクボタンを押すだけ。** 録音を停止すると、そのまま報告書を自動で作成します。")
    auto_mode = st.checkbox(
        "録音を停止したら自動で報告書を作成する", value=True,
        help="オフにすると、下の生成ボタンを押したときだけ作成します。",
    )
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
        audio_source = "record"
        st.audio(recorded, format="audio/wav")
        size_mb = len(recorded) / (1024 * 1024)
        st.caption(f"録音サイズ: {size_mb:.1f} MB")
        if size_mb > MAX_FILE_MB:
            st.error(f"録音が{MAX_FILE_MB}MBを超えました。ファイルを分割してください。")
            audio_bytes = None

with tab_upload:
    uploaded = st.file_uploader(
        f"対応形式: mp3 / m4a / wav / webm / mp4　（上限 {MAX_FILE_MB}MB・最大2時間）",
        type=SUPPORTED_TYPES,
        label_visibility="collapsed",
    )
    if uploaded:
        file_mb = len(uploaded.getvalue()) / (1024 * 1024)
        st.caption(f"ファイル名: {uploaded.name}　｜　サイズ: {file_mb:.1f} MB")
        if file_mb > MAX_FILE_MB:
            st.error(f"ファイルサイズが上限({MAX_FILE_MB}MB)を超えています。2時間以内の音声に分割してください。")
        else:
            audio_bytes = uploaded.getvalue()
            audio_filename = uploaded.name
            audio_source = "upload"

# ─── STEP 2: 個別支援計画書（任意） ──────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 2</b>　個別支援計画書をアップロード（任意・推奨）</div>",
    unsafe_allow_html=True,
)
st.caption("アップロードすると、計画書の目標・課題をもとに報告書の精度が上がります。")

plan_file = st.file_uploader(
    "個別支援計画書（Excel）",
    type=["xlsx", "xls"],
    key="plan_file",
)

plan_text = ""
if plan_file:
    try:
        plan_text = read_plan(plan_file.getvalue())
        st.success(f"✅ 個別支援計画書を読み込みました（{len(plan_text)}文字）")
        with st.expander("読み込み内容を確認する", expanded=False):
            st.text_area("内容", plan_text, height=150, label_visibility="collapsed")
    except Exception as e:
        st.error(f"個別支援計画書の読み込みに失敗しました: {e}")
        plan_text = ""

# ─── STEP 3: 既存ファイルへの追記（任意） ─────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 3</b>　既存ファイルに追記する場合はアップロード（任意）</div>",
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

# ─── STEP 4: 実行ボタン ────────────────────────────────────────
st.markdown(
    "<div class='step-box'><b>STEP 4</b>　報告書を生成</div>",
    unsafe_allow_html=True,
)

run_btn = st.button("🚀 モニタリング報告書を生成する", use_container_width=True, disabled=audio_bytes is None)

# ボタン1つで完結：録音停止後、同じ音声に対して一度だけ自動生成する
audio_hash = hashlib.md5(audio_bytes).hexdigest() if audio_bytes else ""
auto_run = bool(
    audio_bytes
    and audio_source == "record"
    and locals().get("auto_mode", False)
    and audio_hash != st.session_state.get("last_done_hash")
)
if auto_run:
    st.info("🎤 録音を検出しました。自動で報告書を作成しています…")

if (run_btn or auto_run) and audio_bytes:
    st.session_state["last_done_hash"] = audio_hash
    # 文字起こし（長尺は自動分割）
    is_large = len(audio_bytes) > 24 * 1024 * 1024
    spinner_msg = "音声を分割して文字起こし中... (長尺音声は数分かかります)" if is_large else "音声を文字起こし中... (1〜2分かかります)"

    progress_bar = st.progress(0)
    status_text = st.empty()

    def on_progress(current: int, total: int) -> None:
        progress_bar.progress(current / total)
        if total > 1:
            status_text.text(f"チャンク {current} / {total} を処理中...")

    status_text.text(spinner_msg)
    try:
        transcript = transcribe(audio_bytes, audio_filename, progress_callback=on_progress)
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"文字起こしに失敗しました: {e}")
        st.stop()

    progress_bar.empty()
    status_text.empty()
    st.success("✅ 文字起こし完了")

    with st.expander("文字起こし全文を確認する", expanded=False):
        st.text_area("テキスト", transcript, height=200, label_visibility="collapsed")

    # 報告書整形
    with st.spinner("モニタリング報告書を整形中..."):
        try:
            minutes = summarize(transcript, plan_text=plan_text)
        except Exception as e:
            st.error(f"報告書整形に失敗しました: {e}")
            st.stop()

    st.success("✅ 報告書整形完了")

    # ─── プレビュー ────────────────────────────────────────────
    st.divider()
    st.markdown("### 📋 モニタリング報告書 プレビュー")

    a = minutes.get("A_基本情報", {})
    b = minutes.get("B_支援計画", {})
    c = minutes.get("C_モニタリング結果", {})

    # 基本情報サマリー
    col1, col2, col3 = st.columns(3)
    col1.markdown(f"**利用者氏名**: {a.get('利用者氏名') or '—'}")
    col2.markdown(f"**実施日**: {a.get('モニタリング実施日') or '—'}")
    col3.markdown(f"**作成者**: {a.get('作成者氏名') or '—'}")

    achievement = c.get("達成状況の評価", {})
    判定 = achievement.get("判定", "") if isinstance(achievement, dict) else str(achievement)
    badge = {"達成": "🟢", "一部達成": "🟡", "未達成": "🔴"}.get(判定, "⚪")
    変更 = c.get("計画変更の要否", "")
    変更badge = "🔴 要" if 変更 == "要" else "🟢 不要"
    col1.markdown(f"**達成状況**: {badge} {判定}")
    col2.markdown(f"**計画変更**: {変更badge}")
    col3.markdown(f"**次回予定**: {c.get('次回モニタリング予定日') or '—'}")

    st.markdown("---")

    # A. 基本情報
    with st.expander("▶ A. 基本情報", expanded=False):
        st.markdown(f"- 前回実施日: {a.get('前回モニタリング実施日') or '—'}")
        st.markdown(f"- 個別支援計画作成日: {a.get('個別支援計画作成日') or '—'}")

    # B. 支援計画
    with st.expander("▶ B. 支援計画", expanded=False):
        st.markdown(f"**支援の全体方針**: {b.get('支援の全体方針') or '—'}")
        st.markdown(f"**長期目標**: {b.get('長期目標') or '—'}")
        st.markdown(f"**短期目標**: {b.get('短期目標') or '—'}")

    # C. モニタリング結果
    C_FIELDS = [
        ("全体の状況",            c.get("全体の状況", "")),
        ("本人の感想・満足度",    c.get("本人の感想・満足度", "")),
        ("家族・保護者の意向",    c.get("家族・保護者の意向", "")),
        ("達成状況の評価",        f"【{判定}】{achievement.get('詳細','')}" if isinstance(achievement,dict) else str(achievement)),
        ("達成されない原因の分析", c.get("達成されない原因の分析", "")),
        ("今後の対応・支援方針",  c.get("今後の対応・支援方針", "")),
        ("その他留意事項",        c.get("その他留意事項", "")),
    ]
    for label, value in C_FIELDS:
        if value and value not in ("該当なし", "特になし", "面談中言及なし"):
            with st.expander(f"▶ {label}", expanded=False):
                st.write(value)

    # ─── STEP 4: ダウンロード ──────────────────────────────────
    st.divider()
    st.markdown(
        "<div class='step-box'><b>STEP 4</b>　ダウンロード</div>",
        unsafe_allow_html=True,
    )

    name = a.get("利用者氏名", "").replace(" ", "") or "報告書"
    create_date = (a.get("モニタリング実施日") or "").replace("年", "").replace("月", "").replace("日", "")
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
