"""
mail_drafter.py — Yahoo Japan IMAP で36協定書PDFを下書き保存
Yahoo Japan: imap.mail.yahoo.co.jp:993 (SSL)
下書きフォルダ名: "Draft"
"""
import imaplib
import logging
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from urllib.parse import quote

logger = logging.getLogger("yasuda_36kyotei")

IMAP_HOST = "imap.mail.yahoo.co.jp"
IMAP_PORT = 993
DRAFT_FOLDER = "Draft"


def save_draft(
    to_address: str,
    subject: str,
    body: str,
    pdf_bytes: bytes,
    pdf_filename: str,
    imap_user: str,
    imap_password: str,
    from_address: str = "",
) -> dict:
    """Yahoo IMAP経由でPDF添付メールを下書きフォルダに保存する

    Args:
        to_address: 宛先メールアドレス
        subject: 件名
        body: 本文
        pdf_bytes: PDF のバイト列
        pdf_filename: 添付ファイル名（例: "36協定書_○○.pdf"）
        imap_user: Yahoo メールアドレス
        imap_password: Yahoo パスワード
        from_address: 差出人アドレス（省略時は imap_user）

    Returns:
        {"to": str, "subject": str, "status": str}
    """
    # MIMEメッセージ作成
    msg = MIMEMultipart()
    msg["From"] = from_address or imap_user
    msg["To"] = to_address
    msg["Subject"] = subject
    msg["Date"] = formatdate()

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # PDF添付（MIMEBase application/pdf）
    encoded_filename = quote(pdf_filename)
    part = MIMEBase("application", "pdf", name=pdf_filename)
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename*=UTF-8''{encoded_filename}",
    )
    msg.attach(part)

    result = {"to": to_address, "subject": subject, "status": "未保存"}
    try:
        with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=30) as m:
            m.login(imap_user, imap_password)
            internal_date = imaplib.Time2Internaldate(datetime.now(timezone.utc))
            m.append(DRAFT_FOLDER, r"\Draft", internal_date, msg.as_bytes())
        result["status"] = "下書き保存成功"
        logger.info(f"[下書き保存成功] To: {to_address}")
    except imaplib.IMAP4.error as e:
        result["status"] = f"IMAP認証エラー: {str(e)}"
        logger.error(f"[IMAP認証失敗] {e}")
    except Exception as e:
        result["status"] = f"失敗: {str(e)}"
        logger.error(f"[IMAP下書き失敗] To: {to_address} | Error: {e}")

    return result
