"""
ollama_client.py
Ollama API とのやり取りを担う。

Phase 2 構成: 2モデル パイプライン
  Analyst AI (qwen2.5:3b) : 数値データ → 構造化 JSON
  Writer  AI (qwen3:8b)   : 構造化 JSON → 日本語ビジネス文章
"""

import json
import logging
import re
from collections.abc import Callable

import httpx

from config import (
    MODEL_ANALYST,
    MODEL_WRITER,
    OLLAMA_CONNECT_TIMEOUT,
    OLLAMA_GENERATE,
    OLLAMA_GENERATE_TIMEOUT,
    OLLAMA_TAGS_URL,
)

logger = logging.getLogger(__name__)

# 外部から参照できるよう再エクスポート
__all__ = [
    "MODEL_ANALYST",
    "MODEL_WRITER",
    "check_ollama",
    "generate",
    "build_analyst_prompt",
    "parse_analyst_json",
    "build_writer_prompt",
    "parse_writer_response",
]


def check_ollama(model: str) -> None:
    """
    Ollama の起動確認とモデル存在確認を行う。
    問題があれば RuntimeError を即座に送出する（20分タイムアウトを防ぐ事前チェック）。
    """
    try:
        resp = httpx.get(OLLAMA_TAGS_URL, timeout=OLLAMA_CONNECT_TIMEOUT)
        resp.raise_for_status()
        tags = resp.json()
    except httpx.ConnectError:
        raise RuntimeError(
            "Ollama に接続できません。http://localhost:11434 が起動しているか確認してください。"
        )
    except Exception as e:
        raise RuntimeError(f"Ollama の状態確認に失敗しました: {e}")

    # モデル名の先頭部分で前方一致チェック（タグなし/タグあり両対応）
    available = [m.get("name", "") for m in tags.get("models", [])]
    base_name = model.split(":")[0]
    if not any(m == model or m.startswith(base_name + ":") for m in available):
        raise RuntimeError(
            f"モデル '{model}' が Ollama にありません。\n"
            f"利用可能: {available}\n"
            f"`ollama pull {model}` で取得してください。"
        )


# ── 汎用生成関数 ──────────────────────────────────────────────
def generate(
    prompt: str,
    model: str = MODEL_WRITER,
    on_token: Callable[[int], None] | None = None,
) -> str:
    """
    Ollama にプロンプトを送信し、生成テキストを返す。
    stream=True でトークンを逐次受信することで接続が無音のままハングするのを防ぐ。
    on_token(token_count) は生成トークン数が更新されるたびに呼ばれる任意コールバック。
    """
    logger.info(f"Ollama リクエスト送信 (model={model}, prompt_len={len(prompt)})")
    payload = {
        "model":  model,
        "prompt": prompt,
        "stream": True,
        "think":  False,  # qwen3 系の thinking モードを無効化
    }
    try:
        chunks: list[str] = []
        token_count = 0
        with httpx.stream(
            "POST", OLLAMA_GENERATE, json=payload,
            timeout=httpx.Timeout(
                connect=OLLAMA_CONNECT_TIMEOUT,
                read=OLLAMA_GENERATE_TIMEOUT,
                write=30.0,
                pool=5.0,
            ),
        ) as resp:
            resp.raise_for_status()
            for raw_line in resp.iter_lines():
                if not raw_line:
                    continue
                try:
                    chunk = json.loads(raw_line)
                except json.JSONDecodeError:
                    continue
                token = chunk.get("response", "")
                if token:
                    chunks.append(token)
                    token_count += 1
                    if on_token and token_count % 30 == 0:
                        on_token(token_count)
                if chunk.get("done"):
                    break

        text = "".join(chunks)
        # <think>…</think> ブロックが混入している場合は除去する
        text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()
        logger.info(f"Ollama レスポンス受信完了 (tokens={token_count}, len={len(text)})")
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
    except httpx.HTTPError as e:
        logger.error(f"Ollama 通信エラー: {type(e).__name__}: {e}")
        raise RuntimeError(f"Ollama との通信中にエラーが発生しました: {e}")


# ── Analyst AI (qwen2.5:3b) ───────────────────────────────────
def build_analyst_prompt(raw_summary: str) -> str:
    """
    数値データを構造化 JSON に変換するプロンプト。
    小型モデルで確実に JSON のみ出力させるため指示を簡潔にする。
    """
    return (
        "You are a data analyst. Analyze the sales data below and output ONLY a JSON object.\n"
        "No explanations, no markdown code blocks, just the raw JSON.\n\n"
        "Required JSON format:\n"
        "{\n"
        '  "period": "集計期間の文字列",\n'
        '  "total_sales": 総売上金額(数値),\n'
        '  "total_qty": 総販売数量(数値),\n'
        '  "top_products": [{"name": "商品名", "amount": 数値}],\n'
        '  "bottom_products": [{"name": "商品名", "amount": 数値}],\n'
        '  "top_regions": [{"name": "地域名", "amount": 数値}],\n'
        '  "bottom_regions": [{"name": "地域名", "amount": 数値}],\n'
        '  "top_reps": [{"name": "担当者名", "amount": 数値}],\n'
        '  "key_facts": ["重要な数値の事実（日本語）"],\n'
        '  "concerns": ["懸念点・低迷要因（日本語）"],\n'
        '  "yoy_change": {"2024年": "+12.3%", "2023年": "+8.5%"}\n'
        "}\n\n"
        "Sales data:\n"
        f"{raw_summary}"
    )


def parse_analyst_json(response: str) -> dict:
    """
    Analyst AI の出力から JSON を抽出する。
    モデルが余分なテキストを出力した場合も {} ブロックを検索して取り出す。
    """
    # ```json ... ``` ブロックを優先的に抽出
    md_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", response, re.DOTALL)
    if md_match:
        json_str = md_match.group(1)
    else:
        brace_match = re.search(r"\{.*\}", response, re.DOTALL)
        json_str = brace_match.group(0) if brace_match else "{}"

    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        logger.warning(f"Analyst JSON パース失敗。フォールバック使用。response={response[:200]}")
        return {}


# ── Writer AI (qwen3:8b) ──────────────────────────────────────
def build_writer_prompt(
    analyst_data: dict,
    raw_summary: str,
    rag_context: str = "",
    extra_context: str = "",
) -> str:
    """
    Analyst の構造化データを元に日本語ビジネス文章を生成するプロンプト。
    rag_context   : 過去レポートからの参考情報 (RAG)
    extra_context : ユーザーが入力した追加プロンプト
    """
    if analyst_data:
        data_section = f"【分析データ (JSON)】\n{json.dumps(analyst_data, ensure_ascii=False, indent=2)}"
    else:
        logger.warning("Analyst データが空のため raw_summary をフォールバックとして使用")
        data_section = f"【売上データ】\n{raw_summary}"

    rag_section = (
        f"\n【過去レポートからの参考情報】\n"
        f"以下は類似する過去の報告書から抜粋した文脈です。"
        f"傾向の継続・改善状況・前回との差分を意識して文章を構成してください。\n"
        f"{rag_context}\n"
    ) if rag_context else ""

    # extra_context はスタイル指示の直後・出力フォーマットの前に置く
    # → LLM が出力形式を決定する前に読むため、確実に反映される
    extra_section = (
        f"【最優先追加指示】※以下の指示をサマリー・分析の両セクションに必ず反映すること\n"
        f"{extra_context.strip()}\n\n"
    ) if extra_context.strip() else ""

    return (
        "あなたはトップコンサルティングファームのシニアアナリストです。\n"
        "以下の売上分析データを元に、経営陣向けの報告書用テキストを日本語で作成してください。\n\n"
        "【文章スタイル（厳守）】\n"
        "- 各セクション 150〜200字以内、簡潔に要点のみ記述\n"
        "- 箇条書き（・ または 数字+.）を積極的に使う（1セクションあたり 3〜5項目）\n"
        "- 数値は必ず入れる（例: 売上 ¥XX万、前月比 +X%）\n"
        "- 曖昧な表現・冗長な修飾語は排除する\n"
        "- 結論ファースト（最初に結論、次に根拠・数値）\n\n"
        f"{extra_section}"
        "【出力形式】\n"
        "---SUMMARY---\n"
        "（今期の売上サマリーを箇条書きでここに記述）\n"
        "---ANALYSIS---\n"
        "（課題・改善策・次期方針を箇条書きでここに記述）\n\n"
        f"{data_section}\n"
        f"{rag_section}"
    )


def parse_writer_response(response: str) -> tuple[str, str]:
    """Writer AI のレスポンスからサマリーと分析テキストを分離する。"""
    if "---SUMMARY---" in response and "---ANALYSIS---" in response:
        parts         = response.split("---ANALYSIS---")
        summary_text  = parts[0].replace("---SUMMARY---", "").strip()
        analysis_text = parts[1].strip() if len(parts) > 1 else ""
    else:
        logger.warning("Writer レスポンスにセクションマーカーなし。フォールバック分割を使用。")
        mid           = len(response) // 2
        summary_text  = response[:mid].strip()
        analysis_text = response[mid:].strip()

    return summary_text, analysis_text
