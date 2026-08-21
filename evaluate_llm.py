"""
LLM-based RAG evaluation script.
Reads questions from eval_qa.json, calls the RAG retrieval API + Claude Sonnet,
and outputs results to eval_llm_result.csv.

Cost optimizations:
  - Batch API: ask_llm and judge_answer submitted as batches (50% cost reduction)
  - Prompt caching: system prompt and judge instruction cached across requests
"""

import argparse
import csv
import json
import os
import sys
import time
from pathlib import Path

import anthropic
import requests
from dotenv import load_dotenv

load_dotenv()

MODEL = "claude-sonnet-5"
RAG_API_URL = os.getenv("RAG_API_URL", "http://localhost:8000/retrieval")
RAG_API_KEY = os.getenv("RAG_API_KEY", "test-api-key")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
TOP_K = int(os.getenv("TOP_K", "5"))

SYSTEM_PROMPT = """あなたは社内ナレッジベースを参照して質問に答えるアシスタントです。
提供されたコンテキスト情報のみを根拠として、簡潔なテキストで回答してください。
マークダウン記法（見出し・箇条書き・太字など）は使わず、プレーンテキストで答えてください。
コンテキストに情報が含まれていない場合は「提供された情報からは回答できません」とだけ答えてください。"""

JUDGE_INSTRUCTION = (
    "あなたは回答の採点者です。以下の質問に対して、LLMの回答が模範回答の主要な情報を含んでいれば「OK」、"
    "誤りや重要な情報の欠落があれば「NG」とだけ答えてください。\n\n"
)


def retrieve(question: str) -> list[dict]:
    payload = {
        "query": question,
        "retrieval_setting": {"top_k": TOP_K, "score_threshold": 0.0},
    }
    headers = {"Authorization": f"Bearer {RAG_API_KEY}", "Content-Type": "application/json"}
    resp = requests.post(RAG_API_URL, json=payload, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("records", [])


def build_context(records: list[dict]) -> str:
    parts = []
    for i, r in enumerate(records, 1):
        title = r.get("title", "不明")
        content = r.get("content", "").strip()
        parts.append(f"[{i}] {title}\n{content}")
    return "\n\n".join(parts)


def poll_batch(client: anthropic.Anthropic, batch_id: str) -> dict[str, str]:
    """Wait for batch to complete and return {custom_id: text} results."""
    while True:
        batch = client.beta.messages.batches.retrieve(batch_id)
        counts = batch.request_counts
        print(f"  [{batch.processing_status}] processing={counts.processing} "
              f"succeeded={counts.succeeded} errored={counts.errored}")
        if batch.processing_status == "ended":
            break
        time.sleep(30)

    results = {}
    for result in client.beta.messages.batches.results(batch_id):
        if result.result.type == "succeeded":
            # ThinkingBlockが先頭に来る場合があるため、TextBlockを明示的に探す
            text_block = next(
                (b for b in result.result.message.content if b.type == "text"), None
            )
            results[result.custom_id] = text_block.text if text_block else ""
        else:
            results[result.custom_id] = f"[BATCH ERROR] {result.result.error}"
    return results


def main():
    parser = argparse.ArgumentParser(description="LLMベースのRAG評価スクリプト")
    parser.add_argument("--collection", default=None,
                        help="評価対象のChromaDBコレクション名（APIサーバー側で設定）")
    parser.add_argument("--output", default=None,
                        help="結果CSVの出力先 (default: eval_llm_result.csv)")
    parser.add_argument("--qa-file", default="eval_qa.json",
                        help="評価Q&AファイルのパスまたはファイルQA名 (default: eval_qa.json)")
    args = parser.parse_args()

    if not ANTHROPIC_API_KEY:
        print("Error: ANTHROPIC_API_KEY is not set in .env", file=sys.stderr)
        sys.exit(1)

    qa_path = Path(__file__).parent / args.qa_file
    questions = json.loads(qa_path.read_text(encoding="utf-8"))

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # ── Phase 1: RAG Retrieval (sequential, local API) ──────────────────────
    print("=== Phase 1: RAG Retrieval ===")
    retrieved: list[list[dict]] = []
    for q in questions:
        print(f"  Q{q['id']}: retrieving...", end=" ", flush=True)
        try:
            records = retrieve(q["question"])
            retrieved.append(records)
            top_score = records[0].get("score", 0.0) if records else 0.0
            print(f"score={top_score:.3f}")
        except Exception as e:
            print(f"[ERROR] {e}")
            retrieved.append([])

    # ── Phase 2: Batch ask_llm ───────────────────────────────────────────────
    print("\n=== Phase 2: Batch ask_llm ===")
    ask_requests = []
    for q, records in zip(questions, retrieved):
        context = build_context(records)
        user_message = (
            "以下のコンテキスト情報を参照して質問に答えてください。\n\n"
            f"【コンテキスト】\n{context}\n\n"
            f"【質問】\n{q['question']}"
        )
        ask_requests.append({
            "custom_id": f"ask-{q['id']}",
            "params": {
                "model": MODEL,
                "max_tokens": 1024,
                # cache_control on system prompt: shared across all requests → cache hit from 2nd onward
                "system": [{"type": "text", "text": SYSTEM_PROMPT,
                             "cache_control": {"type": "ephemeral"}}],
                "messages": [{"role": "user", "content": user_message}],
            },
        })

    ask_batch = client.beta.messages.batches.create(requests=ask_requests)
    print(f"  Submitted: {ask_batch.id} ({len(ask_requests)} requests)")
    ask_results = poll_batch(client, ask_batch.id)

    # ── Phase 3: Batch judge_answer ──────────────────────────────────────────
    print("\n=== Phase 3: Batch judge_answer ===")
    judge_requests = []
    for q in questions:
        qid = q["id"]
        llm_answer = ask_results.get(f"ask-{qid}", "[NO RESULT]")
        judge_requests.append({
            "custom_id": f"judge-{qid}",
            "params": {
                "model": MODEL,
                "max_tokens": 8,
                "messages": [{
                    "role": "user",
                    "content": [
                        # cache_control on fixed instruction prefix
                        {"type": "text", "text": JUDGE_INSTRUCTION,
                         "cache_control": {"type": "ephemeral"}},
                        {"type": "text", "text": (
                            f"【質問】{q['question']}\n\n"
                            f"【模範回答】{q['expected_answer']}\n\n"
                            f"【LLMの回答】{llm_answer}"
                        )},
                    ],
                }],
            },
        })

    judge_batch = client.beta.messages.batches.create(requests=judge_requests)
    print(f"  Submitted: {judge_batch.id} ({len(judge_requests)} requests)")
    judge_results = poll_batch(client, judge_batch.id)

    # ── Phase 4: Write CSV ───────────────────────────────────────────────────
    output_filename = args.output or "eval_llm_result.csv"
    output_path = Path(__file__).parent / output_filename
    fieldnames = [
        "id", "question", "expected_answer", "llm_answer", "eval_result",
        "currently_answerable", "retrieved_titles", "top_score", "source_hit",
    ]

    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for q, records in zip(questions, retrieved):
            qid = q["id"]
            titles = [r.get("title", "") for r in records]
            top_score = records[0].get("score", 0.0) if records else 0.0
            expected_source = q.get("source", "")
            source_hit = any(expected_source in t for t in titles) if expected_source else False

            llm_answer = ask_results.get(f"ask-{qid}", "[NO RESULT]")
            verdict = judge_results.get(f"judge-{qid}", "NG")
            eval_result = "OK" if verdict.strip().startswith("OK") else "NG"

            writer.writerow({
                "id": qid,
                "question": q["question"],
                "expected_answer": q["expected_answer"],
                "llm_answer": llm_answer,
                "eval_result": eval_result,
                "currently_answerable": q.get("currently_answerable", True),
                "retrieved_titles": " | ".join(titles),
                "top_score": f"{top_score:.4f}",
                "source_hit": source_hit,
            })
            print(f"Q{qid}: {eval_result} (score={top_score:.3f}, hit={source_hit})")

    print(f"\nResults saved to {output_path}")


if __name__ == "__main__":
    main()
