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

import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../backend"))

from main import app  # noqa: E402

client = TestClient(app, raise_server_exceptions=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "../../data/template.pptx")


def _make_excel_bytes() -> bytes:
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


def _make_csv_bytes() -> bytes:
    data = {
        "日付":     ["2025-01-10", "2025-02-15"],
        "商品名":   ["商品A", "商品B"],
        "担当者":   ["田中", "佐藤"],
        "地域":     ["東京", "大阪"],
        "数量":     [2, 1],
        "売上金額": [100_000, 80_000],
    }
    df = pd.DataFrame(data)
    return df.to_csv(index=False).encode("utf-8")


def _make_pptx_bytes() -> bytes:
    with open(TEMPLATE_PATH, "rb") as f:
        return f.read()


# ── テスト ──────────────────────────────────────────────────

def test_health_endpoint():
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json() == {"status": "ok"}


@patch("routers.report.generate", return_value="---SUMMARY---\nMock サマリー\n---ANALYSIS---\nMock 分析")
def test_generate_success_excel(mock_generate):
    """正常系(Excel): SSE ストリームで done イベントが返ることを確認。"""
    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("sales.xlsx", _make_excel_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 200
    assert "text/event-stream" in res.headers["content-type"]
    assert '"done": true' in res.text or '"done":true' in res.text
    assert mock_generate.call_count == 1


@patch("routers.report.generate", return_value="---SUMMARY---\nMock サマリー\n---ANALYSIS---\nMock 分析")
def test_generate_success_csv(mock_generate):
    """正常系(CSV): SSE ストリームで done イベントが返ることを確認。"""
    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("sales.csv", _make_csv_bytes(), "text/csv"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 200
    assert '"done": true' in res.text or '"done":true' in res.text


@patch("routers.report.generate", side_effect=RuntimeError("Ollama に接続できません"))
def test_generate_ollama_error(mock_generate):
    """Ollama 接続エラー時に SSE error イベントが返ることを確認。"""
    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("sales.xlsx", _make_excel_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 200
    assert '"error"' in res.text
    assert "Ollama" in res.text


def test_generate_invalid_excel():
    """不正なファイルを送ると SSE error イベントが返ることを確認。"""
    res = client.post(
        "/api/generate",
        files={
            "excel_file":    ("bad.xlsx", b"not an excel", "application/octet-stream"),
            "template_file": ("template.pptx", _make_pptx_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation"),
        },
    )
    assert res.status_code == 200
    assert '"error"' in res.text


def test_download_not_found():
    """レポート未生成時に 404 が返ることを確認。"""
    import shutil
    output = os.path.join(os.path.dirname(__file__), "../../output/report.pptx")
    backup = output + ".bak"
    if os.path.exists(output):
        shutil.move(output, backup)
    try:
        res = client.get("/api/download")
        assert res.status_code == 404
    finally:
        if os.path.exists(backup):
            shutil.move(backup, output)
