"""
routers/references.py
過去レポート PPTX を RAG 用にベクター登録・管理するエンドポイント。

POST   /api/references/upload  — PPTX をアップロードして ChromaDB に登録
GET    /api/references          — 登録済みファイル一覧
DELETE /api/references/{file_id} — 指定ファイルを削除
"""

import logging
import os
import tempfile

from fastapi import APIRouter, File, HTTPException, UploadFile

from services.rag_store import delete_report, list_registered, register_report

logger = logging.getLogger(__name__)
router = APIRouter()


@router.post("/api/references/upload")
async def upload_reference(file: UploadFile = File(..., description="過去レポート PPTX")):
    """過去レポート PPTX を受け取り ChromaDB に登録する。"""
    if not (file.filename or "").lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="PPTX ファイルのみ登録できます。")

    data = await file.read()
    filename = file.filename or "report.pptx"

    tmpdir = tempfile.mkdtemp()
    pptx_path = os.path.join(tmpdir, filename)
    try:
        with open(pptx_path, "wb") as f:
            f.write(data)
        chunk_count = register_report(pptx_path, filename)
    except RuntimeError as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)

    return {"filename": filename, "chunks": chunk_count, "status": "registered"}


@router.get("/api/references")
async def get_references():
    """登録済みの過去レポート一覧を返す。"""
    return {"references": list_registered()}


@router.delete("/api/references/{file_id}")
async def remove_reference(file_id: str):
    """指定した file_id の過去レポートを ChromaDB から削除する。"""
    deleted = delete_report(file_id)
    if deleted == 0:
        raise HTTPException(status_code=404, detail="指定されたファイルが見つかりません。")
    return {"file_id": file_id, "deleted_chunks": deleted}
