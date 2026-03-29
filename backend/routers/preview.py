"""
routers/preview.py
POST /api/preview  — アップロードされた Excel/CSV の先頭 5 行と必須列チェック結果を返す。
"""

import logging
import os
import tempfile

import pandas as pd
from fastapi import APIRouter, File, HTTPException, UploadFile

logger = logging.getLogger(__name__)
router = APIRouter()

REQUIRED_COLS = {"日付", "商品名", "担当者", "地域", "数量", "売上金額"}
PREVIEW_ROWS  = 5


@router.post("/api/preview")
async def preview_file(file: UploadFile = File(..., description="Excel または CSV ファイル")):
    """
    ファイルの先頭 PREVIEW_ROWS 行をテーブル形式で返し、必須列の過不足もチェックする。
    レスポンス:
        columns     : 列名リスト
        rows        : 先頭 N 行のデータ（各行は値リスト）
        missing_cols: 不足している必須列名リスト
    """
    filename = file.filename or ""
    suffix = os.path.splitext(filename)[1].lower()
    if suffix not in (".xlsx", ".xls", ".csv"):
        raise HTTPException(
            status_code=422,
            detail="対応していないファイル形式です。.xlsx / .xls / .csv のみ対応しています。",
        )

    data = await file.read()

    tmpdir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmpdir, f"preview{suffix}")
    try:
        with open(tmp_path, "wb") as f:
            f.write(data)

        try:
            if suffix == ".csv":
                try:
                    df = pd.read_csv(tmp_path, encoding="utf-8", nrows=PREVIEW_ROWS + 1)
                except UnicodeDecodeError:
                    df = pd.read_csv(tmp_path, encoding="shift-jis", nrows=PREVIEW_ROWS + 1)
            else:
                df = pd.read_excel(tmp_path, nrows=PREVIEW_ROWS)
        except Exception as e:
            raise HTTPException(status_code=422, detail=f"ファイルの読み込みに失敗しました: {e}")

        columns = list(df.columns)
        rows = [
            [str(v) if pd.notna(v) else "" for v in row]
            for row in df.head(PREVIEW_ROWS).values.tolist()
        ]
        missing_cols = sorted(REQUIRED_COLS - set(columns))

        return {
            "columns":      columns,
            "rows":         rows,
            "missing_cols": missing_cols,
        }
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)
