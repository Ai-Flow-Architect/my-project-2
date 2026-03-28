"""
36協定自動化ツール - Streamlit Webアプリ
社労士事務所向け。Excelアップロード → Word協定書生成 → ZIPダウンロード
"""
import io
import os
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

# --- パス設定（同ディレクトリのモジュールをインポート）---
sys.path.insert(0, str(Path(__file__).parent))
from excel_reader import read_excel
from word_generator import generate_word, FORM_NAMES

# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="36協定自動化ツール",
    page_icon="📄",
    layout="centered",
)

# ============================================================
# カスタムCSS（非エンジニア向け・見やすいUI）
# ============================================================
st.markdown("""
<style>
    .main-title {
        font-size: 1.8rem;
        font-weight: bold;
        color: #1a1a2e;
        text-align: center;
        padding: 1rem 0 0.3rem 0;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #555;
        text-align: center;
        margin-bottom: 2rem;
    }
    .step-box {
        background: #f8f9ff;
        border-left: 4px solid #4a6cf7;
        padding: 1rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 1.2rem 0 0.5rem 0;
    }
    .step-label {
        font-size: 0.75rem;
        font-weight: bold;
        color: #4a6cf7;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .step-title {
        font-size: 1.05rem;
        font-weight: bold;
        color: #1a1a2e;
        margin-top: 0.2rem;
    }
    .result-card {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin: 0.5rem 0;
    }
    .error-card {
        background: #fef2f2;
        border: 1px solid #fca5a5;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin: 0.5rem 0;
    }
    .footer {
        text-align: center;
        color: #aaa;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# パスワード認証
# ============================================================
def check_password() -> bool:
    """パスワード認証画面。正しければ True を返す。"""

    # Streamlit Cloud の st.secrets["password"] または環境変数 APP_PASSWORD を使用
    correct_pw = None
    try:
        correct_pw = st.secrets["password"]
    except Exception:
        correct_pw = os.environ.get("APP_PASSWORD", "")

    if not correct_pw:
        # パスワード未設定の場合はそのまま通す（開発時）
        return True

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    # --- パスワード入力画面 ---
    st.markdown('<div class="main-title">📄 36協定自動化ツール</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">朝日事務所</div>', unsafe_allow_html=True)
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("#### 🔒 パスワードを入力してください")
        pw = st.text_input("パスワード", type="password", key="pw_input", label_visibility="collapsed",
                           placeholder="パスワードを入力")
        if st.button("ログイン", use_container_width=True, type="primary"):
            if pw == correct_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("パスワードが違います。もう一度お試しください。")
    return False


# ============================================================
# Word生成 → ZIPバイナリを返す
# ============================================================
def generate_zip(records: list[dict]) -> tuple[bytes, list[dict]]:
    """
    全レコードの Word ファイルを生成し、ZIP バイナリとして返す。
    Returns:
        (zip_bytes, results) results は [{name, form, status, error}, ...]
    """
    results = []
    zip_buffer = io.BytesIO()

    with tempfile.TemporaryDirectory() as tmpdir, zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, record in enumerate(records):
            name = record.get("事業所名", f"企業{i+1}")
            form_type = record.get("様式パターン", "9")
            form_label = FORM_NAMES.get(form_type, form_type)

            try:
                out_path = generate_word(record, output_dir=tmpdir)
                filename = Path(out_path).name
                zf.write(out_path, arcname=filename)
                results.append({"事業所名": name, "様式": form_label, "結果": "✅ 生成完了", "エラー": ""})
            except Exception as e:
                results.append({"事業所名": name, "様式": form_label, "結果": "❌ 失敗", "エラー": str(e)})

    return zip_buffer.getvalue(), results


# ============================================================
# メインアプリ
# ============================================================
def main() -> None:
    st.markdown('<div class="main-title">📄 36協定自動化ツール</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Excelをアップロードするだけで、36協定書（Word）を一括生成します</div>',
                unsafe_allow_html=True)

    # --- STEP 1: Excel アップロード ---
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 1</div>
        <div class="step-title">📂 Excelファイルをアップロード</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "36協定データの Excel ファイル（.xlsx）を選択してください",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    if uploaded is None:
        st.info("👆 Excelファイルを選択すると、処理が始まります。")
        _show_footer()
        return

    # --- Excel 読み取り ---
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
        records = read_excel(tmp_path)
    except Exception as e:
        st.markdown(f'<div class="error-card">❌ Excelの読み取りに失敗しました。<br><small>{e}</small></div>',
                    unsafe_allow_html=True)
        return
    finally:
        os.unlink(tmp_path)

    if not records:
        st.warning("Excelにデータが見つかりませんでした。内容を確認してください。")
        return

    # --- STEP 2: プレビュー ---
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 2</div>
        <div class="step-title">📋 読み取り結果の確認</div>
    </div>
    """, unsafe_allow_html=True)

    st.success(f"**{len(records)} 件** のデータを読み取りました。")

    preview_rows = []
    for i, r in enumerate(records):
        form_type = r.get("様式パターン", "9")
        form_label = FORM_NAMES.get(form_type, form_type)
        preview_rows.append({
            "#": i + 1,
            "事業所名": r.get("事業所名", "（未入力）"),
            "事業主名": r.get("事業主名", "（未入力）"),
            "様式": form_label,
            "特別条項": "あり" if form_type in ("9_2", "9_3", "9_4", "9_5") else "なし",
        })

    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    # --- STEP 3: Word生成 & ダウンロード ---
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 3</div>
        <div class="step-title">📝 Word協定書を一括生成してダウンロード</div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⚡ Word協定書を生成する", type="primary", use_container_width=True):
        with st.spinner("協定書を生成しています…"):
            zip_bytes, results = generate_zip(records)

        # 結果表示
        success_count = sum(1 for r in results if "✅" in r["結果"])
        fail_count = len(results) - success_count

        if fail_count == 0:
            st.markdown(f'<div class="result-card">✅ <strong>{success_count} 件</strong> すべて生成完了しました。</div>',
                        unsafe_allow_html=True)
        else:
            st.warning(f"{success_count} 件成功 / {fail_count} 件失敗")

        # 詳細テーブル
        st.dataframe(results, use_container_width=True, hide_index=True)

        # ZIPダウンロード
        if success_count > 0:
            st.download_button(
                label="📥 協定書をまとめてダウンロード（ZIP）",
                data=zip_bytes,
                file_name="36協定書_一括.zip",
                mime="application/zip",
                use_container_width=True,
            )

    _show_footer()


def _show_footer():
    st.markdown(
        '<div class="footer">36協定自動化ツール｜朝日事務所</div>',
        unsafe_allow_html=True,
    )


# ============================================================
# エントリーポイント
# ============================================================
if check_password():
    main()
