"""
routers/references.py
過去レポート PPTX を RAG 用にベクター登録・管理するエンドポイント。

POST   /api/references/upload          — PPTX をアップロードして ChromaDB に登録
GET    /api/references                 — 登録済みファイル一覧
GET    /api/references/{file_id}/chunks — 登録済みファイルのチャンク一覧
DELETE /api/references/{file_id}       — 指定ファイルを削除
"""

import logging
import os
import tempfile

from fastapi import APIRouter, File, HTTPException, UploadFile

from services.rag_store import delete_report, get_chunks_for_file, list_registered, register_report

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


@router.get("/api/references/{file_id}/chunks")
async def get_reference_chunks(file_id: str):
    """登録済みファイルから抽出したチャンク（参考情報）一覧を返す。"""
    chunks = get_chunks_for_file(file_id)
    if chunks is None:
        raise HTTPException(status_code=404, detail="指定されたファイルが見つかりません。")
    return {"file_id": file_id, "chunks": chunks}


@router.delete("/api/references/{file_id}")
async def remove_reference(file_id: str):
    """指定した file_id の過去レポートを ChromaDB から削除する。"""
    deleted = delete_report(file_id)
    if deleted == 0:
        raise HTTPException(status_code=404, detail="指定されたファイルが見つかりません。")
    return {"file_id": file_id, "deleted_chunks": deleted}
