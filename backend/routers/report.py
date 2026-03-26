"""
routers/report.py
POST /api/generate  — Excel + テンプレートから PPTX を生成してダウンロードさせる。
"""

import logging
import os
import tempfile
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from services.excel_reader import read_and_summarize
from services.ollama_client import build_combined_prompt, generate, parse_combined_response
from services.pptx_generator import generate_pptx

logger = logging.getLogger(__name__)
router = APIRouter()

OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


@router.post("/api/generate")
async def generate_report(
    excel_file:    UploadFile = File(..., description="売上データ Excel (.xlsx)"),
    template_file: UploadFile = File(..., description="PPTX テンプレート (.pptx)"),
):
    """
    Excel と PPTX テンプレートを受け取り、Ollama で分析して報告書 PPTX を返す。
    CPU 推論のため処理時間が長くなる場合があります（最大 6 分）。
    """
    logger.info(f"generate_report 開始: excel={excel_file.filename}, template={template_file.filename}")

    # ── 一時ファイルに保存 ──────────────────────────────
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path    = os.path.join(tmpdir, "sales_data.xlsx")
        template_path = os.path.join(tmpdir, "template.pptx")
        output_path   = str(OUTPUT_DIR / "report.pptx")

        with open(excel_path, "wb") as f:
            f.write(await excel_file.read())
        with open(template_path, "wb") as f:
            f.write(await template_file.read())

        # ── Excel 集計 ────────────────────────────────────
        try:
            summary_data = read_and_summarize(excel_path)
        except (FileNotFoundError, ValueError) as e:
            logger.error(f"Excel 読み込みエラー: {e}")
            raise HTTPException(status_code=400, detail=str(e))

        # ── Ollama でテキスト生成（1回で summary + analysis を取得）────
        try:
            logger.info("Ollama: summary + analysis 生成中 (1リクエスト)…")
            combined = generate(build_combined_prompt(summary_data["raw_summary"]))
            summary_text, analysis_text = parse_combined_response(combined)
            logger.info(f"Ollama 生成完了 summary={len(summary_text)}字 analysis={len(analysis_text)}字")
        except RuntimeError as e:
            logger.error(f"Ollama エラー: {e}")
            raise HTTPException(status_code=503, detail=str(e))

        # ── PPTX 生成 ─────────────────────────────────────
        try:
            generate_pptx(
                template_path=template_path,
                output_path=output_path,
                summary_text=summary_text,
                analysis_text=analysis_text,
                period=summary_data["period"],
            )
        except (FileNotFoundError, Exception) as e:
            logger.error(f"PPTX 生成エラー: {e}")
            raise HTTPException(status_code=500, detail=f"PPTX 生成に失敗しました: {e}")

    logger.info("generate_report 完了 — ファイル返却")
    return FileResponse(
        path=output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="report.pptx",
    )
