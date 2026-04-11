"""
transcriber.py — Whisper API で音声ファイルを文字起こし
"""
from __future__ import annotations

import os
import tempfile
from pathlib import Path

from openai import OpenAI


def transcribe(audio_bytes: bytes, filename: str) -> str:
    """
    音声バイト列を受け取り、文字起こしテキストを返す。
    対応フォーマット: mp3 / m4a / wav / webm
    上限: 25MB（Whisper API 制限）
    """
    api_key = os.environ.get("OPENAI_API_KEY", "")
    client = OpenAI(api_key=api_key)

    suffix = Path(filename).suffix.lower() or ".mp3"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(audio_bytes)
        tmp_path = tmp.name

    try:
        with open(tmp_path, "rb") as f:
            response = client.audio.transcriptions.create(
                model="whisper-1",
                file=f,
                language="ja",
                response_format="text",
            )
        return str(response)
    finally:
        Path(tmp_path).unlink(missing_ok=True)
