"""
test_rag.py — RAG 関連のユニットテスト
ChromaDB と Ollama embed は Mock 化してテストする。
"""
import os
import sys
from unittest.mock import MagicMock, patch

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "../../backend"))


# ── rag_store のユニットテスト ────────────────────────────────

class TestRagStore:
    """rag_store.py の各関数をモックを使ってテスト。"""

    def _make_collection_mock(self):
        col = MagicMock()
        col.count.return_value = 0
        col.get.return_value = {"ids": [], "metadatas": []}
        return col

    @patch("services.rag_store._get_collection")
    @patch("services.rag_store.embed_text", return_value=[0.1] * 768)
    @patch("services.rag_store.extract_chunks_from_pptx", return_value=["売上は好調です。前月比10%増。"])
    def test_register_report(self, mock_extract, mock_embed, mock_get_col):
        col = self._make_collection_mock()
        mock_get_col.return_value = col

        from services.rag_store import register_report
        count = register_report("/fake/path.pptx", "report_2024.pptx")

        assert count == 1
        col.add.assert_called_once()
        args = col.add.call_args.kwargs
        assert args["documents"] == ["売上は好調です。前月比10%増。"]
        assert len(args["embeddings"][0]) == 768

    @patch("services.rag_store._get_collection")
    @patch("services.rag_store.embed_text", return_value=[0.2] * 768)
    def test_search_context_with_data(self, mock_embed, mock_get_col):
        col = self._make_collection_mock()
        col.count.return_value = 1
        col.query.return_value = {
            "documents": [["過去の売上サマリーテキスト"]],
            "metadatas": [[{"filename": "report.pptx", "file_id": "abc123", "chunk_idx": 0}]],
            "distances": [[0.15]],   # similarity = 0.85
        }
        mock_get_col.return_value = col

        from services.rag_store import search_context
        ctx = search_context("今月の売上データ")

        assert "過去の売上サマリーテキスト" in ctx
        assert "report.pptx" in ctx

    @patch("services.rag_store._get_collection")
    def test_search_context_empty(self, mock_get_col):
        col = self._make_collection_mock()
        col.count.return_value = 0
        mock_get_col.return_value = col

        from services.rag_store import search_context
        ctx = search_context("クエリ")
        assert ctx == ""

    @patch("services.rag_store._get_collection")
    @patch("services.rag_store.embed_text", return_value=[0.1] * 768)
    def test_search_context_low_similarity_filtered(self, mock_embed, mock_get_col):
        """類似度が低いチャンクは除外される。"""
        col = self._make_collection_mock()
        col.count.return_value = 1
        col.query.return_value = {
            "documents": [["関係ないテキスト"]],
            "metadatas": [[{"filename": "x.pptx", "file_id": "x", "chunk_idx": 0}]],
            "distances": [[0.85]],   # similarity = 0.15 → 閾値 0.3 以下
        }
        mock_get_col.return_value = col

        from services.rag_store import search_context
        ctx = search_context("全然違うトピック")
        assert ctx == ""

    @patch("services.rag_store._get_collection")
    def test_list_registered(self, mock_get_col):
        col = self._make_collection_mock()
        col.count.return_value = 2
        col.get.return_value = {
            "metadatas": [
                {"filename": "a.pptx", "file_id": "id_a", "chunk_idx": 0},
                {"filename": "a.pptx", "file_id": "id_a", "chunk_idx": 1},
            ]
        }
        mock_get_col.return_value = col

        from services.rag_store import list_registered
        result = list_registered()
        assert len(result) == 1
        assert result[0]["filename"] == "a.pptx"
        assert result[0]["chunks"] == 2

    @patch("services.rag_store._get_collection")
    def test_delete_report(self, mock_get_col):
        col = self._make_collection_mock()
        col.get.return_value = {"ids": ["id_a_0", "id_a_1"]}
        mock_get_col.return_value = col

        from services.rag_store import delete_report
        deleted = delete_report("id_a")
        assert deleted == 2
        col.delete.assert_called_once_with(ids=["id_a_0", "id_a_1"])


# ── references エンドポイントのテスト ────────────────────────

class TestReferencesAPI:

    def _client(self):
        from fastapi.testclient import TestClient
        from main import app
        return TestClient(app)

    @patch("routers.references.register_report", return_value=3)
    def test_upload_reference(self, mock_reg):
        client = self._client()
        res = client.post(
            "/api/references/upload",
            files={"file": ("report_2024.pptx", b"fake pptx data", "application/octet-stream")},
        )
        assert res.status_code == 200
        data = res.json()
        assert data["chunks"] == 3
        assert data["status"] == "registered"

    def test_upload_non_pptx_rejected(self):
        client = self._client()
        res = client.post(
            "/api/references/upload",
            files={"file": ("data.xlsx", b"fake", "application/octet-stream")},
        )
        assert res.status_code == 400

    @patch("routers.references.list_registered", return_value=[
        {"filename": "r.pptx", "file_id": "abc", "chunks": 2}
    ])
    def test_get_references(self, mock_list):
        client = self._client()
        res = client.get("/api/references")
        assert res.status_code == 200
        assert len(res.json()["references"]) == 1

    @patch("routers.references.delete_report", return_value=2)
    def test_delete_reference(self, mock_del):
        client = self._client()
        res = client.delete("/api/references/abc123")
        assert res.status_code == 200
        assert res.json()["deleted_chunks"] == 2

    @patch("routers.references.delete_report", return_value=0)
    def test_delete_not_found(self, mock_del):
        client = self._client()
        res = client.delete("/api/references/nonexistent")
        assert res.status_code == 404


# ── ollama_client の build_writer_prompt RAG 対応テスト ───────

def test_writer_prompt_includes_rag_context():
    from services.ollama_client import build_writer_prompt
    prompt = build_writer_prompt(
        {"total_sales": 100},
        "raw data",
        rag_context="過去レポートの文脈テキスト",
    )
    assert "過去レポートからの参考情報" in prompt
    assert "過去レポートの文脈テキスト" in prompt


def test_writer_prompt_no_rag_context():
    from services.ollama_client import build_writer_prompt
    prompt = build_writer_prompt({"total_sales": 100}, "raw data", rag_context="")
    assert "過去レポートからの参考情報" not in prompt
