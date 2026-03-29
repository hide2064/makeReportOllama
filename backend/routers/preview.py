"""
routers/preview.py
POST /api/preview  — アップロードされた Excel/CSV の先頭 5 行と必須列チェック結果、
                     およびチャート用集計データを返す。
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
    ファイルの先頭 PREVIEW_ROWS 行と必須列チェック結果、チャート用集計データを返す。

    レスポンス:
        columns     : 列名リスト
        rows        : 先頭 N 行のデータ（各行は値リスト）
        missing_cols: 不足している必須列名リスト
        chart_data  : 集計データ（必須列が揃っている場合のみ）/ null
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

        # 全行読み込み（プレビューとチャート集計を兼用）
        try:
            if suffix == ".csv":
                try:
                    df = pd.read_csv(tmp_path, encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(tmp_path, encoding="shift-jis")
            else:
                df = pd.read_excel(tmp_path)
        except Exception as e:
            raise HTTPException(status_code=422, detail=f"ファイルの読み込みに失敗しました: {e}")

        columns      = list(df.columns)
        rows         = [
            [str(v) if pd.notna(v) else "" for v in row]
            for row in df.head(PREVIEW_ROWS).values.tolist()
        ]
        missing_cols = sorted(REQUIRED_COLS - set(columns))

        # チャートデータ集計（必須列が揃っている場合のみ）
        chart_data = None
        if not missing_cols:
            try:
                df["日付"] = pd.to_datetime(df["日付"], errors="coerce")
                df = df.dropna(subset=["日付"])

                if not df.empty:
                    product_totals: dict[str, int] = (
                        df.groupby("商品名")["売上金額"].sum()
                        .sort_values(ascending=False)
                        .head(8)
                        .astype(int)
                        .to_dict()
                    )
                    df["月"] = df["日付"].dt.to_period("M").astype(str)
                    monthly_totals: dict[str, int] = (
                        df.groupby("月")["売上金額"].sum()
                        .sort_index()
                        .tail(6)
                        .astype(int)
                        .to_dict()
                    )
                    chart_data = {
                        "product_totals": product_totals,
                        "monthly_totals": monthly_totals,
                        "total_amount":   int(df["売上金額"].sum()),
                        "total_qty":      int(df["数量"].sum()),
                        "period": (
                            f"{df['日付'].min().strftime('%Y/%m/%d')} 〜 "
                            f"{df['日付'].max().strftime('%Y/%m/%d')}"
                        ),
                    }
            except Exception as e:
                logger.warning(f"チャートデータ集計に失敗: {e}")

        return {
            "columns":      columns,
            "rows":         rows,
            "missing_cols": missing_cols,
            "chart_data":   chart_data,
        }
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)
