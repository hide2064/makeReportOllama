"""
routers/templates.py
GET /api/templates — data/ ディレクトリ内の template*.pptx 一覧を返す。
"""
from pathlib import Path

from fastapi import APIRouter

router = APIRouter()

DATA_DIR = Path(__file__).parent.parent.parent / "data"


@router.get("/api/templates")
async def list_templates():
    """data/ ディレクトリ内の template*.pptx ファイル一覧を返す。"""
    templates = [
        {"name": p.name}
        for p in sorted(DATA_DIR.glob("template*.pptx"))
        if p.is_file()
    ]
    return {"templates": templates}
