"""
routers/report.py
POST /api/generate  — SSE でステップ進捗を送りながら PPTX を生成する。
GET  /api/download  — 生成済み PPTX をダウンロードさせる。
"""

import json
import logging
import os
import tempfile
from pathlib import Path

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse

from services.excel_reader import read_and_summarize
from services.ollama_client import build_combined_prompt, generate, parse_combined_response
from services.pptx_generator import generate_pptx

logger = logging.getLogger(__name__)
router = APIRouter()

OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)


def _sse(data: dict) -> str:
    """SSE フォーマット 1 イベント分を返す。"""
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"


@router.post("/api/generate")
async def generate_report(
    excel_file:    UploadFile = File(..., description="売上データ Excel/CSV"),
    template_file: UploadFile = File(..., description="PPTX テンプレート (.pptx)"),
):
    """
    Excel/CSV と PPTX テンプレートを受け取り、進捗を SSE で送信しながら報告書 PPTX を生成する。
    完了後は GET /api/download でファイルを取得すること。
    """
    logger.info(f"generate_report 開始: excel={excel_file.filename}, template={template_file.filename}")
    excel_data    = await excel_file.read()
    template_data = await template_file.read()
    excel_filename = excel_file.filename or "sales_data.xlsx"

    async def event_stream():
        with tempfile.TemporaryDirectory() as tmpdir:
            suffix        = Path(excel_filename).suffix.lower() or ".xlsx"
            excel_path    = os.path.join(tmpdir, f"sales_data{suffix}")
            template_path = os.path.join(tmpdir, "template.pptx")
            output_path   = str(OUTPUT_DIR / "report.pptx")

            with open(excel_path, "wb") as f:
                f.write(excel_data)
            with open(template_path, "wb") as f:
                f.write(template_data)

            # ── Step 1: Excel 読み込み ──────────────────────
            yield _sse({"step": "[1/3]  Excel / CSV を読み込んでいます..."})
            logger.info("Step 1: Excel 読み込み")
            try:
                summary_data = read_and_summarize(excel_path)
            except (FileNotFoundError, ValueError) as e:
                logger.error(f"Excel 読み込みエラー: {e}")
                yield _sse({"error": str(e)})
                return

            # ── Step 2: Ollama 推論 ─────────────────────────
            yield _sse({"step": "[2/3]  Ollama (ローカル LLM) で売上を分析中です...\n"
                                "       CPU 推論のため数分かかります。このまましばらくお待ちください。"})
            logger.info("Step 2: Ollama 推論")
            try:
                combined = generate(build_combined_prompt(summary_data["raw_summary"]))
                summary_text, analysis_text = parse_combined_response(combined)
                logger.info(f"Ollama 完了 summary={len(summary_text)}字 analysis={len(analysis_text)}字")
            except RuntimeError as e:
                logger.error(f"Ollama エラー: {e}")
                yield _sse({"error": str(e)})
                return

            # ── Step 3: PPTX 生成 ──────────────────────────
            yield _sse({"step": "[3/3]  PowerPoint レポートを生成しています..."})
            logger.info("Step 3: PPTX 生成")
            try:
                generate_pptx(
                    template_path=template_path,
                    output_path=output_path,
                    summary_text=summary_text,
                    analysis_text=analysis_text,
                    period=summary_data["period"],
                )
            except Exception as e:
                logger.error(f"PPTX 生成エラー: {e}")
                yield _sse({"error": f"PPTX 生成に失敗しました: {e}"})
                return

            logger.info("generate_report 完了")
            yield _sse({"done": True})

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@router.get("/api/download")
async def download_report():
    """生成済み report.pptx をダウンロードさせる。"""
    output_path = OUTPUT_DIR / "report.pptx"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="レポートファイルが見つかりません。先に生成してください。")
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="report.pptx",
    )
