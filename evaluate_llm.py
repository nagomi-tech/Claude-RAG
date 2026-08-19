"""
LLM-based RAG evaluation script.
Reads questions from eval_qa.json, calls the RAG retrieval API + Claude Sonnet,
and outputs results to eval_llm_result.csv.
"""

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

RAG_API_URL = os.getenv("RAG_API_URL", "http://localhost:8000/retrieval")
RAG_API_KEY = os.getenv("RAG_API_KEY", "test-api-key")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
TOP_K = int(os.getenv("TOP_K", "5"))

SYSTEM_PROMPT = """あなたは社内ナレッジベースを参照して質問に答えるアシスタントです。
提供されたコンテキスト情報のみを根拠として、簡潔なテキストで回答してください。
マークダウン記法（見出し・箇条書き・太字など）は使わず、プレーンテキストで答えてください。
コンテキストに情報が含まれていない場合は「提供された情報からは回答できません」とだけ答えてください。"""


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


def judge_answer(client: anthropic.Anthropic, question: str, expected: str, llm_answer: str) -> str:
    """Return 'OK' if llm_answer is sufficiently correct, 'NG' otherwise."""
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8,
        messages=[{
            "role": "user",
            "content": (
                "あなたは回答の採点者です。以下の質問に対して、LLMの回答が模範回答の主要な情報を含んでいれば「OK」、"
                "誤りや重要な情報の欠落があれば「NG」とだけ答えてください。\n\n"
                f"【質問】{question}\n\n"
                f"【模範回答】{expected}\n\n"
                f"【LLMの回答】{llm_answer}"
            ),
        }],
    )
    verdict = message.content[0].text.strip()
    return "OK" if verdict.startswith("OK") else "NG"


def ask_llm(client: anthropic.Anthropic, question: str, context: str) -> str:
    user_message = f"""以下のコンテキスト情報を参照して質問に答えてください。

【コンテキスト】
{context}

【質問】
{question}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        system=SYSTEM_PROMPT,

        
        messages=[{"role": "user", "content": user_message}],
    )
    return message.content[0].text


def main():
    if not ANTHROPIC_API_KEY:
        print("Error: ANTHROPIC_API_KEY is not set in .env", file=sys.stderr)
        sys.exit(1)

    qa_path = Path(__file__).parent / "eval_qa.json"
    questions = json.loads(qa_path.read_text(encoding="utf-8"))

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    output_path = Path(__file__).parent / "eval_llm_result.csv"
    fieldnames = [
        "id",
        "question",
        "expected_answer",
        "llm_answer",
        "eval_result",
        "currently_answerable",
        "retrieved_titles",
        "top_score",
        "source_hit",
    ]

    with output_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for q in questions:
            qid = q["id"]
            question = q["question"]
            expected = q["expected_answer"]
            expected_source = q.get("source", "")
            currently_answerable = q.get("currently_answerable", True)

            print(f"Q{qid}: {question[:50]}...", end=" ", flush=True)

            try:
                records = retrieve(question)
            except Exception as e:
                print(f"[RETRIEVAL ERROR] {e}")
                writer.writerow({
                    "id": qid,
                    "question": question,
                    "expected_answer": expected,
                    "llm_answer": f"[RETRIEVAL ERROR] {e}",
                    "eval_result": "NG",
                    "currently_answerable": currently_answerable,
                    "retrieved_titles": "",
                    "top_score": "",
                    "source_hit": False,
                })
                continue

            titles = [r.get("title", "") for r in records]
            top_score = records[0].get("score", 0.0) if records else 0.0
            source_hit = any(expected_source in t for t in titles) if expected_source else False

            context = build_context(records)

            try:
                llm_answer = ask_llm(client, question, context)
            except Exception as e:
                print(f"[LLM ERROR] {e}")
                llm_answer = f"[LLM ERROR] {e}"

            try:
                eval_result = judge_answer(client, question, expected, llm_answer)
            except Exception as e:
                eval_result = "NG"

            writer.writerow({
                "id": qid,
                "question": question,
                "expected_answer": expected,
                "llm_answer": llm_answer,
                "eval_result": eval_result,
                "currently_answerable": currently_answerable,
                "retrieved_titles": " | ".join(titles),
                "top_score": f"{top_score:.4f}",
                "source_hit": source_hit,
            })
            print(f"{eval_result} (score={top_score:.3f}, hit={source_hit})")

            # Rate limit: ~1 req/sec
            time.sleep(1.0)

    print(f"\nResults saved to {output_path}")


if __name__ == "__main__":
    main()
