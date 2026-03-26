"""
ollama_client.py
Ollama API (http://localhost:11434) とのやり取りを担う。
CPU推論のため、タイムアウトは 5 分以上に設定する。
"""

import logging

import httpx

logger = logging.getLogger(__name__)

OLLAMA_URL     = "http://localhost:11434/api/generate"
REQUEST_TIMEOUT = 1200  # seconds (20分)
DEFAULT_MODEL  = "qwen3-vl:8b"  # Ollama にインストール済みのモデル名


def generate(prompt: str, model: str = DEFAULT_MODEL) -> str:
    """
    Ollama にプロンプトを送信し、生成テキストを返す。
    stream=False でレスポンスをまとめて受け取る。
    """
    logger.info(f"Ollama リクエスト送信 (model={model}, prompt_len={len(prompt)})")
    payload = {
        "model":  model,
        "prompt": prompt,
        "stream": False,
    }
    try:
        resp = httpx.post(OLLAMA_URL, json=payload, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        text = resp.json().get("response", "")
        logger.info(f"Ollama レスポンス受信 (len={len(text)})")
        return text
    except httpx.TimeoutException:
        logger.error("Ollama タイムアウト")
        raise RuntimeError("Ollama がタイムアウトしました。モデルが起動しているか確認してください。")
    except httpx.HTTPStatusError as e:
        logger.error(f"Ollama HTTP エラー: {e.response.status_code} {e.response.text}")
        raise RuntimeError(f"Ollama エラー: {e.response.status_code}")
    except httpx.ConnectError:
        logger.error("Ollama に接続できません")
        raise RuntimeError("Ollama に接続できません。http://localhost:11434 が起動しているか確認してください。")


def build_combined_prompt(raw_summary: str) -> str:
    """サマリーと分析を1回のリクエストで生成するプロンプト。"""
    return (
        "あなたは優秀なビジネスアナリストです。\n"
        "以下の売上データを分析し、下記の2つのセクションを日本語で作成してください。\n"
        "各セクションは300字程度、箇条書きを使わず文章形式で記述してください。\n\n"
        "【出力形式】\n"
        "---SUMMARY---\n"
        "（売上サマリーをここに記述）\n"
        "---ANALYSIS---\n"
        "（課題・所見と来月の改善策・方針をここに記述）\n\n"
        "【売上データ】\n"
        f"{raw_summary}"
    )


def parse_combined_response(response: str) -> tuple[str, str]:
    """1回のレスポンスからサマリーと分析テキストを分離する。"""
    summary_text  = ""
    analysis_text = ""

    if "---SUMMARY---" in response and "---ANALYSIS---" in response:
        parts = response.split("---ANALYSIS---")
        summary_part  = parts[0].replace("---SUMMARY---", "").strip()
        analysis_part = parts[1].strip() if len(parts) > 1 else ""
        summary_text  = summary_part
        analysis_text = analysis_part
    else:
        # フォールバック: レスポンス全体をサマリーに使用
        mid = len(response) // 2
        summary_text  = response[:mid].strip()
        analysis_text = response[mid:].strip()

    return summary_text, analysis_text


def build_summary_prompt(raw_summary: str) -> str:
    return (
        "あなたは優秀なビジネスアナリストです。\n"
        "以下の売上データを分析し、経営陣向けの売上サマリーを日本語で300字程度で作成してください。\n"
        "箇条書きを使わず、文章形式で記述してください。\n\n"
        f"{raw_summary}"
    )


def build_analysis_prompt(raw_summary: str) -> str:
    return (
        "あなたは優秀なビジネスアナリストです。\n"
        "以下の売上データを元に、課題・所見と来月に向けた改善策・方針を日本語で300字程度で作成してください。\n"
        "箇条書きを使わず、文章形式で記述してください。\n\n"
        f"{raw_summary}"
    )
