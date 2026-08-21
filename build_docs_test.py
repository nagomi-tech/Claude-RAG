"""
docs_test/ をv2.1方式（全ページVision）でインデックスするスクリプト。

v2.1方式:
  - テキスト: PDF/DOCX/PPTX/XLSXをloaders経由で抽出 → chunk分割
  - Vision: 全PDFの全ページをレンダリングしてClaude Visionへ送信（フィルタなし）

Cost optimizations:
  - Batch API: 全ページのVision処理をまとめて送信（50%コスト削減）

Usage:
  python build_docs_test.py
  python build_docs_test.py --collection docs_test_v21  # デフォルト
  python build_docs_test.py --force                      # 既存チャンクを削除して再構築
"""

import argparse
import base64
import os
import time
from pathlib import Path

import anthropic
import pymupdf
from dotenv import load_dotenv
from langchain_chroma import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_core.documents import Document
from langchain_text_splitters import RecursiveCharacterTextSplitter

# ingest.py のローダーを流用
from ingest import load_pdf, load_docx, load_pptx, load_xlsx, EMBED_MODEL

load_dotenv()

MODEL = "claude-sonnet-5"
DOCS_TEST = Path(__file__).parent / "docs_test"
CHROMA_DIR = Path(__file__).parent / "chroma_db"

LOADERS = {
    ".pdf": load_pdf,
    ".docx": load_docx,
    ".pptx": load_pptx,
    ".xlsx": load_xlsx,
}

VISION_PROMPT = (
    "これは「{doc_name}」のp.{page}です。\n"
    "この画像に含まれるすべての情報（数値、名前、項目名、構造、関係性など）を"
    "漏れなく日本語のテキストとして書き起こしてください。\n"
    "グラフの場合は各項目の値を、表の場合は行と列の対応を正確に、"
    "フロー図の場合は各ステップと条件分岐を具体的に記述してください。"
)


def build_text_chunks() -> list[Document]:
    """docs_test/ 以下のファイルをテキスト抽出してチャンク化する。"""
    raw_docs = []
    for path in sorted(DOCS_TEST.iterdir()):
        ext = path.suffix.lower()
        loader = LOADERS.get(ext)
        if loader is None:
            continue
        print(f"  テキスト読み込み: {path.name}")
        try:
            docs = loader(path)
            raw_docs.extend(docs)
        except Exception as e:
            print(f"    [エラー] {e}")

    splitter = RecursiveCharacterTextSplitter(
        chunk_size=500,
        chunk_overlap=100,
        separators=["\n\n", "\n", "。", "、", " ", ""],
    )

    chunks = []
    for doc in raw_docs:
        if doc.metadata.get("type") == "table":
            chunks.append(doc)
        else:
            chunks.extend(splitter.split_documents([doc]))

    return chunks


def build_vision_chunks(client: anthropic.Anthropic) -> list[Document]:
    """全PDFの全ページをバッチAPIでVision処理する（v2.1方式）。"""
    pdf_files = sorted(f for f in DOCS_TEST.iterdir() if f.suffix.lower() == ".pdf")

    # ── 全ページの画像データを収集 ──────────────────────────────────────────
    page_data: list[tuple] = []  # (path, page_idx, total, b64)
    for path in pdf_files:
        print(f"  レンダリング: {path.name}")
        doc = pymupdf.open(str(path))
        total = len(doc)
        for i in range(total):
            page = doc[i]
            mat = pymupdf.Matrix(1.5, 1.5)  # 150% 解像度
            pix = page.get_pixmap(matrix=mat)
            b64 = base64.standard_b64encode(pix.tobytes("png")).decode()
            page_data.append((path, i, total, b64))

    print(f"  合計 {len(page_data)} ページをバッチ送信...")

    # ── バッチリクエスト構築 ─────────────────────────────────────────────────
    batch_requests = []
    for path, i, total, b64 in page_data:
        custom_id = f"vision-{path.stem}-p{i + 1}"
        batch_requests.append({
            "custom_id": custom_id,
            "params": {
                "model": MODEL,
                "max_tokens": 2048,
                "messages": [{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {"type": "base64", "media_type": "image/png", "data": b64},
                        },
                        {
                            "type": "text",
                            "text": VISION_PROMPT.format(doc_name=path.name, page=i + 1),
                        },
                    ],
                }],
            },
        })

    # ── バッチ送信 ───────────────────────────────────────────────────────────
    batch = client.beta.messages.batches.create(requests=batch_requests)
    print(f"  バッチID: {batch.id}")

    # ── 完了待機 ─────────────────────────────────────────────────────────────
    while True:
        batch = client.beta.messages.batches.retrieve(batch.id)
        counts = batch.request_counts
        print(f"  [{batch.processing_status}] processing={counts.processing} "
              f"succeeded={counts.succeeded} errored={counts.errored}")
        if batch.processing_status == "ended":
            break
        time.sleep(30)

    # ── 結果収集 ─────────────────────────────────────────────────────────────
    result_map: dict = {}
    for result in client.beta.messages.batches.results(batch.id):
        if result.result.type == "succeeded":
            result_map[result.custom_id] = result.result.message.content[0].text

    # ── Document リスト構築（順序保持） ──────────────────────────────────────
    vision_docs = []
    for path, i, total, _ in page_data:
        custom_id = f"vision-{path.stem}-p{i + 1}"
        text = result_map.get(custom_id)
        if text:
            vision_docs.append(Document(
                page_content=text,
                metadata={
                    "source": path.name,
                    "file_type": "vision_extracted",
                    "page": i + 1,
                },
            ))
            print(f"    {path.name} p.{i + 1}/{total}: {len(text)}文字")
        else:
            print(f"    {path.name} p.{i + 1}/{total}: [エラー]")

    return vision_docs


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--collection", default="docs_test_v21")
    parser.add_argument("--force", action="store_true",
                        help="既存コレクションを削除して再構築")
    args = parser.parse_args()

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise SystemExit("Error: ANTHROPIC_API_KEY not set")

    embeddings = HuggingFaceEmbeddings(
        model_name=EMBED_MODEL,
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )

    if args.force:
        # 既存コレクションを削除
        vs = Chroma(
            persist_directory=str(CHROMA_DIR),
            embedding_function=embeddings,
            collection_name=args.collection,
        )
        vs.delete_collection()
        print(f"既存コレクション '{args.collection}' を削除しました。\n")

    client = anthropic.Anthropic(api_key=api_key)

    print("=== テキスト抽出 ===")
    text_chunks = build_text_chunks()
    table_count = sum(1 for d in text_chunks if d.metadata.get("type") == "table")
    print(f"  テキストチャンク: {len(text_chunks) - table_count}件 / テーブル: {table_count}件\n")

    print("=== Vision処理（全ページ・バッチAPI） ===")
    vision_chunks = build_vision_chunks(client)
    print(f"  Visionチャンク: {len(vision_chunks)}件\n")

    all_chunks = text_chunks + vision_chunks
    print(f"=== ChromaDB に保存: collection={args.collection} ===")
    vs = Chroma.from_documents(
        documents=all_chunks,
        embedding=embeddings,
        persist_directory=str(CHROMA_DIR),
        collection_name=args.collection,
    )
    print(f"完了: {vs._collection.count()} チャンクを登録しました。")
    print(f"\nAPIサーバーを以下で起動してください:")
    print(f"  COLLECTION={args.collection} python3 api.py")


if __name__ == "__main__":
    main()
