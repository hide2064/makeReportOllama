"""
routers/models.py
GET /api/models  — ローカル Ollama で利用可能なモデル一覧と現在の設定を返す。
"""

import logging

import httpx
from fastapi import APIRouter

from services.ollama_client import MODEL_ANALYST, MODEL_WRITER

logger = logging.getLogger(__name__)
router = APIRouter()

OLLAMA_TAGS_URL = "http://localhost:11434/api/tags"


@router.get("/api/models")
async def get_models():
    """Ollama で利用可能なモデル一覧と現在の Analyst/Writer 設定を返す。"""
    available: list[str] = []
    try:
        resp = httpx.get(OLLAMA_TAGS_URL, timeout=5.0)
        resp.raise_for_status()
        data = resp.json()
        available = [m["name"] for m in data.get("models", [])]
    except Exception as e:
        logger.warning(f"Ollama モデル一覧取得失敗: {e}")

    return {
        "available": available,
        "current_analyst": MODEL_ANALYST,
        "current_writer": MODEL_WRITER,
    }
