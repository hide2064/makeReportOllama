"""
test_api.py — FastAPI エンドポイントの結合テスト
Ollama 通信部分は Mock 化してテストする。
"""
import io
import os
from unittest.mock import patch

import pandas as pd
import pytest
from fastapi.testclient import TestClient
from pptx import Presentation
from pptx.util import Inches

# main.py をインポート
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../backend"))

from main import app  # noqa: E402

client = TestClient(app, raise_server_exceptions=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "../../data/template.pptx")


def _make_excel_bytes() -> bytes:
    """テスト用 Excel バイト列を生成。"""
    data = {
        "日付":     ["2025-01-10", "2025-02-15"],
        "商品名":   ["商品A", "商品B"],
        "担当者":   ["田中", "佐藤"],
        "地域":     ["東京", "大阪"],
        "数量":     [2, 1],
        "売上金額": [100_000, 80_000],
    }
    df = pd.DataFrame(data)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def _make_pptx_bytes() -> bytes:
    """テスト用 PPTX バイト列を生成（実テンプレートを使用）。"""
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()


# ── テスト ────────────────────────────────────────────────────

def test_health_endpoint():
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json() == {"status": "ok"}


@patch("routers.report.generate", return_value="Mock 生成テキスト")
def test_generate_success(mock_generate):
    """正常系: Ollama を Mock 化して PPTX が返ることを確認。"""
    excel_bytes    = _make_excel_bytes()
    template_bytes = _make_pptx_bytes()

    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("sales.xlsx", excel_bytes,    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "template_file": ("template.pptx", template_bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 200
    assert res.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    # Ollama が 2 回呼ばれること（summary + analysis）
    assert mock_generate.call_count == 2


@patch("routers.report.generate", side_effect=RuntimeError("Ollama に接続できません"))
def test_generate_ollama_error(mock_generate):
    """Ollama 接続エラー時に 503 が返ることを確認。"""
    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("sales.xlsx", _make_excel_bytes(),    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 503


def test_generate_invalid_excel():
    """不正な Excel（空バイト）を送ると 400 が返ることを確認。"""
    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("bad.xlsx", b"not an excel", "application/octet-stream"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 400
