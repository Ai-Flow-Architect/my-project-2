"""
transcriber.py — Whisper API で音声ファイルを文字起こし
25MB超の長尺音声（最大2時間）は自動分割して処理
"""
from __future__ import annotations

import io
import os
import shutil
import tempfile
from pathlib import Path
from typing import Callable

from openai import OpenAI

CHUNK_MINUTES = 25
OVERLAP_SECONDS = 10
MAX_SINGLE_BYTES = 24 * 1024 * 1024
WHISPER_TIMEOUT = 180  # 秒


def _require_api_key() -> str:
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        raise EnvironmentError("OPENAI_API_KEY が設定されていません。環境変数を確認してください。")
    return api_key


def _require_ffmpeg() -> None:
    if shutil.which("ffmpeg") is None:
        raise EnvironmentError(
            "ffmpeg が見つかりません。\n"
            "Ubuntu: sudo apt install ffmpeg\n"
            "Mac:    brew install ffmpeg\n"
            "インストール後に再起動してください。"
        )


def _transcribe_single(client: OpenAI, audio_bytes: bytes, suffix: str) -> str:
    """単一音声チャンクをWhisperで文字起こし"""
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
                timeout=WHISPER_TIMEOUT,
            )
        return str(response)
    finally:
        Path(tmp_path).unlink(missing_ok=True)


def _split_audio(audio_bytes: bytes, filename: str) -> list[bytes]:
    """
    音声をCHUNK_MINUTES分のチャンクに分割する。
    各チャンク末尾にOVERLAP_SECONDS秒のオーバーラップを持たせて文脈切れを防ぐ。
    戻り値: MP3バイト列のリスト
    """
    _require_ffmpeg()
    from pydub import AudioSegment

    chunk_ms = CHUNK_MINUTES * 60 * 1000
    overlap_ms = OVERLAP_SECONDS * 1000
    step_ms = chunk_ms - overlap_ms
    if step_ms <= 0:
        raise ValueError(
            f"OVERLAP_SECONDS ({OVERLAP_SECONDS}) が CHUNK_MINUTES*60 ({CHUNK_MINUTES * 60}) 以上です。"
        )

    suffix = Path(filename).suffix.lower().lstrip(".") or "mp3"
    audio = AudioSegment.from_file(io.BytesIO(audio_bytes), format=suffix)
    total_ms = len(audio)

    chunks: list[bytes] = []
    start_ms = 0
    while start_ms < total_ms:
        end_ms = min(start_ms + chunk_ms, total_ms)
        buf = io.BytesIO()
        audio[start_ms:end_ms].export(buf, format="mp3")
        chunks.append(buf.getvalue())
        if end_ms >= total_ms:
            break
        start_ms = end_ms - overlap_ms

    return chunks


def transcribe(
    audio_bytes: bytes,
    filename: str,
    progress_callback: Callable[[int, int], None] | None = None,
) -> str:
    """
    音声バイト列を文字起こしして返す。
    - 24MB以下: Whisperへ直接送信
    - 24MB超: 25分チャンクに自動分割して順次送信し結合

    progress_callback(current, total) で処理中チャンク番号を通知する。
    """
    api_key = _require_api_key()
    client = OpenAI(api_key=api_key)
    suffix = Path(filename).suffix.lower() or ".mp3"

    if len(audio_bytes) <= MAX_SINGLE_BYTES:
        if progress_callback:
            progress_callback(1, 1)
        return _transcribe_single(client, audio_bytes, suffix)

    chunks = _split_audio(audio_bytes, filename)
    results: list[str] = []
    for i, chunk in enumerate(chunks):
        if progress_callback:
            progress_callback(i + 1, len(chunks))
        text = _transcribe_single(client, chunk, ".mp3")
        results.append(text.strip())

    return "\n\n".join(results)
