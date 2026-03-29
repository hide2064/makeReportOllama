"""
routers/history.py
GET  /api/history             — 直近 N 件の生成履歴を返す。
GET  /api/history/{job_id}/download  — 指定ジョブの PPTX を再ダウンロード。
"""

import logging
from pathlib import Path

from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import FileResponse

from services.history_store import get_history_item, list_history

logger = logging.getLogger(__name__)
router = APIRouter()


@router.get("/api/history")
async def get_history(n: int = Query(default=20, ge=1, le=50)):
    """直近 n 件の生成済みレポート一覧を返す。"""
    return {"history": list_history(n)}


@router.get("/api/history/{job_id}/download")
async def download_history(job_id: str):
    """指定した job_id の PPTX ファイルをダウンロードさせる。"""
    item = get_history_item(job_id)
    if not item:
        raise HTTPException(status_code=404, detail="指定されたジョブが履歴に見つかりません。")

    output_path = item.get("output_path", "")
    if not output_path or not Path(output_path).exists():
        raise HTTPException(status_code=404, detail="ファイルが存在しません。削除された可能性があります。")

    filename = f"report_{job_id}.pptx"
    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )
