"""
36協定自動化ツール - Streamlit Webアプリ
社労士事務所向け。Excelアップロード → Word協定書生成 → メール送信
"""
import io
import os
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from excel_reader import read_excel
from word_generator import generate_word, FORM_NAMES
from mail_sender import create_email, send_email, build_email_body, build_subject

# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="36協定自動化ツール",
    page_icon="📄",
    layout="centered",
)

# ============================================================
# カスタムCSS
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
    correct_pw = None
    try:
        correct_pw = st.secrets["password"]
    except Exception:
        correct_pw = os.environ.get("APP_PASSWORD", "")

    if not correct_pw:
        return True

    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown('<div class="main-title">📄 36協定自動化ツール</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">朝日事務所</div>', unsafe_allow_html=True)
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("#### 🔒 パスワードを入力してください")
        pw = st.text_input("パスワード", type="password", key="pw_input",
                           label_visibility="collapsed", placeholder="パスワードを入力")
        if st.button("ログイン", use_container_width=True, type="primary"):
            if pw == correct_pw:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("パスワードが違います。もう一度お試しください。")
    return False


# ============================================================
# SMTP設定をSecretsから取得
# ============================================================
def get_smtp_config() -> dict:
    keys = ["smtp_server", "smtp_port", "smtp_user", "smtp_password",
            "from_address", "差出人名", "差出人所属", "差出人電話"]
    config = {}
    for k in keys:
        try:
            config[k] = st.secrets[k]
        except Exception:
            config[k] = os.environ.get(k.upper(), "")

    if not config.get("smtp_server"):
        config["smtp_server"] = "smtp.gmail.com"
    if not config.get("smtp_port"):
        config["smtp_port"] = 587

    return config


# ============================================================
# Word生成 → ZIPバイナリ＋ファイルbytes辞書を返す
# ============================================================
def generate_files(records: list[dict]) -> tuple[bytes, dict[str, bytes], list[dict]]:
    results = []
    zip_buffer = io.BytesIO()
    file_bytes: dict[str, bytes] = {}

    with tempfile.TemporaryDirectory() as tmpdir, \
         zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, record in enumerate(records):
            name = record.get("事業所名", f"企業{i+1}")
            form_type = record.get("様式パターン", "9")
            form_label = FORM_NAMES.get(form_type, form_type)

            try:
                out_path = generate_word(record, output_dir=tmpdir)
                filename = Path(out_path).name
                with open(out_path, "rb") as f:
                    file_bytes[filename] = f.read()
                zf.write(out_path, arcname=filename)
                results.append({"事業所名": name, "様式": form_label, "結果": "✅ 生成完了", "エラー": ""})
            except Exception as e:
                results.append({"事業所名": name, "様式": form_label, "結果": "❌ 失敗", "エラー": str(e)})

    return zip_buffer.getvalue(), file_bytes, results


# ============================================================
# メール一括送信
# ============================================================
def send_all_emails(records: list[dict], file_bytes: dict[str, bytes], smtp_config: dict) -> list[dict]:
    results = []
    with tempfile.TemporaryDirectory() as tmpdir:
        for record in records:
            name = record.get("事業所名", "")
            email_addr = record.get("メールアドレス", "")
            form_type = record.get("様式パターン", "9")
            form_label = FORM_NAMES.get(form_type, form_type)
            filename = f"36協定書_{name}_{form_label}.docx"

            if not email_addr:
                results.append({"事業所名": name, "宛先": "（未設定）", "結果": "⚠️ メールアドレスなし"})
                continue

            attachment_path = None
            if filename in file_bytes:
                tmp_path = Path(tmpdir) / filename
                tmp_path.write_bytes(file_bytes[filename])
                attachment_path = str(tmp_path)

            subject = build_subject(record)
            body = build_email_body(record, smtp_config)
            msg = create_email(
                to_address=email_addr,
                subject=subject,
                body=body,
                attachment_path=attachment_path,
                from_address=smtp_config.get("from_address", ""),
            )
            send_result = send_email(
                msg,
                smtp_server=smtp_config.get("smtp_server", "smtp.gmail.com"),
                smtp_port=int(smtp_config.get("smtp_port", 587)),
                username=smtp_config.get("smtp_user", ""),
                password=smtp_config.get("smtp_password", ""),
                dry_run=False,
            )
            results.append({"事業所名": name, "宛先": email_addr, "結果": send_result["status"]})

    return results


# ============================================================
# メインアプリ
# ============================================================
def main() -> None:
    st.markdown('<div class="main-title">📄 36協定自動化ツール</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Excelをアップロードするだけで、36協定書（Word）の生成・メール送信が完了します</div>',
                unsafe_allow_html=True)

    # session_state 初期化
    if "generated_files" not in st.session_state:
        st.session_state.generated_files = {}
    if "zip_bytes" not in st.session_state:
        st.session_state.zip_bytes = None
    if "gen_results" not in st.session_state:
        st.session_state.gen_results = []
    if "records" not in st.session_state:
        st.session_state.records = []
    if "send_results" not in st.session_state:
        st.session_state.send_results = []

    # --------------------------------------------------------
    # STEP 1: Excel アップロード
    # --------------------------------------------------------
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

    # ファイルが差し替わったらsession_stateをリセット
    if "last_filename" not in st.session_state or st.session_state.last_filename != uploaded.name:
        st.session_state.last_filename = uploaded.name
        st.session_state.generated_files = {}
        st.session_state.zip_bytes = None
        st.session_state.gen_results = []
        st.session_state.records = []
        st.session_state.send_results = []

    # Excel 読み取り
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

    st.session_state.records = records

    # --------------------------------------------------------
    # STEP 2: プレビュー
    # --------------------------------------------------------
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
            "メールアドレス": r.get("メールアドレス", "（未設定）"),
        })

    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # STEP 3: Word生成 & ダウンロード
    # --------------------------------------------------------
    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 3</div>
        <div class="step-title">📝 Word協定書を一括生成してダウンロード</div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⚡ Word協定書を生成する", type="primary", use_container_width=True):
        with st.spinner("協定書を生成しています…"):
            zip_bytes, file_bytes, gen_results = generate_files(records)
        st.session_state.zip_bytes = zip_bytes
        st.session_state.generated_files = file_bytes
        st.session_state.gen_results = gen_results
        st.session_state.send_results = []  # 再生成時はメール結果をリセット

    if st.session_state.gen_results:
        success_count = sum(1 for r in st.session_state.gen_results if "✅" in r["結果"])
        fail_count = len(st.session_state.gen_results) - success_count

        if fail_count == 0:
            st.markdown(f'<div class="result-card">✅ <strong>{success_count} 件</strong> すべて生成完了しました。</div>',
                        unsafe_allow_html=True)
        else:
            st.warning(f"{success_count} 件成功 / {fail_count} 件失敗")

        st.dataframe(st.session_state.gen_results, use_container_width=True, hide_index=True)

        if st.session_state.zip_bytes and success_count > 0:
            st.download_button(
                label="📥 協定書をまとめてダウンロード（ZIP）",
                data=st.session_state.zip_bytes,
                file_name="36協定書_一括.zip",
                mime="application/zip",
                use_container_width=True,
            )

    # --------------------------------------------------------
    # STEP 4: メール送信（Word生成完了後のみ表示）
    # --------------------------------------------------------
    if not st.session_state.generated_files:
        _show_footer()
        return

    st.markdown("""
    <div class="step-box">
        <div class="step-label">STEP 4</div>
        <div class="step-title">✉️ 協定書をメールで一括送信</div>
    </div>
    """, unsafe_allow_html=True)

    # 送信予定テーブル
    send_preview = []
    for r in records:
        name = r.get("事業所名", "")
        email_addr = r.get("メールアドレス", "")
        subject = build_subject(r)
        send_preview.append({
            "事業所名": name,
            "宛先メール": email_addr if email_addr else "⚠️ 未設定",
            "件名": subject,
        })

    st.markdown("**送信予定の一覧**")
    st.dataframe(send_preview, use_container_width=True, hide_index=True)

    # SMTP設定確認
    smtp_config = get_smtp_config()
    smtp_ok = bool(smtp_config.get("smtp_user") and smtp_config.get("smtp_password"))

    if not smtp_ok:
        st.error("⚠️ メール送信の設定が完了していません。管理者に連絡してください。（SMTP設定未反映）")
        _show_footer()
        return

    st.info(f"📤 送信元: **{smtp_config.get('from_address', smtp_config.get('smtp_user', ''))}**")

    confirmed = st.checkbox("上記の宛先にメールを送信することを確認しました")

    if confirmed:
        if st.button("📨 メールを一括送信する", type="primary", use_container_width=True):
            with st.spinner("メールを送信しています…"):
                send_results = send_all_emails(
                    st.session_state.records,
                    st.session_state.generated_files,
                    smtp_config,
                )
            st.session_state.send_results = send_results

    if st.session_state.send_results:
        mail_ok = sum(1 for r in st.session_state.send_results if "成功" in r["結果"])
        mail_fail = len(st.session_state.send_results) - mail_ok

        if mail_fail == 0:
            st.markdown(f'<div class="result-card">✅ <strong>{mail_ok} 件</strong> すべて送信完了しました。</div>',
                        unsafe_allow_html=True)
        else:
            st.warning(f"{mail_ok} 件送信成功 / {mail_fail} 件失敗")

        st.dataframe(st.session_state.send_results, use_container_width=True, hide_index=True)

    _show_footer()


def _show_footer():
    st.markdown(
        '<div class="footer">36協定自動化ツール｜朝日事務所</div>',
        unsafe_allow_html=True,
    )


if check_password():
    main()
