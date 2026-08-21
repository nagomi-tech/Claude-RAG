"""
Dify 外部ナレッジベース API 仕様に準拠した FastAPI サーバー。

Dify の期待するリクエスト/レスポンス形式:
  POST /retrieval
  Authorization: Bearer <API_KEY>

  Request:
    {
      "knowledge_id": "...",
      "query": "検索クエリ",
      "retrieval_setting": {
        "top_k": 5,
        "score_threshold": 0.5
      }
    }

  Response:
    {
      "records": [
        {
          "content": "...",
          "score": 0.95,
          "title": "ファイル名",
          "metadata": { "source": "...", "file_type": "..." }
        }
      ]
    }
"""

import os
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel

from langchain_chroma import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings

# --- 設定 ---
API_KEY = os.environ.get("RAG_API_KEY", "dify-rag-secret")
CHROMA_DIR = Path(__file__).parent / "chroma_db"
EMBED_MODEL = "intfloat/multilingual-e5-small"
COLLECTION = os.environ.get("COLLECTION", "dify_rag")

# --- アプリ初期化 ---
app = FastAPI(title="Dify RAG API")

# 埋め込みモデルと ChromaDB をサーバー起動時に一度だけロード
embeddings = HuggingFaceEmbeddings(
    model_name=EMBED_MODEL,
    model_kwargs={"device": "cpu"},
    encode_kwargs={"normalize_embeddings": True},
)

vectorstore = Chroma(
    persist_directory=str(CHROMA_DIR),
    embedding_function=embeddings,
    collection_name=COLLECTION,
)


# --- スキーマ ---
class RetrievalSetting(BaseModel):
    top_k: int = 5
    score_threshold: float = 0.5


class RetrievalRequest(BaseModel):
    knowledge_id: Optional[str] = None
    query: str
    retrieval_setting: RetrievalSetting = RetrievalSetting()


class Record(BaseModel):
    content: str
    score: float
    title: str
    metadata: dict


class RetrievalResponse(BaseModel):
    records: list[Record]


# --- 認証ミドルウェア ---
@app.middleware("http")
async def verify_api_key(request: Request, call_next):
    if request.url.path in ("/", "/health"):
        return await call_next(request)
    auth = request.headers.get("Authorization", "")
    if not auth.startswith("Bearer ") or auth[7:] != API_KEY:
        return JSONResponse(status_code=403, content={"error": "Unauthorized"})
    return await call_next(request)


# --- エンドポイント ---
@app.get("/")
def root():
    return {"status": "ok", "message": "Dify RAG API is running"}


@app.get("/health")
def health():
    try:
        count = len(vectorstore.get()["ids"])
    except Exception:
        count = -1
    return {"status": "ok", "indexed_chunks": count, "collection": COLLECTION}


@app.post("/retrieval", response_model=RetrievalResponse)
def retrieval(req: RetrievalRequest):
    setting = req.retrieval_setting
    results = vectorstore.similarity_search_with_relevance_scores(
        query=req.query,
        k=setting.top_k,
    )

    records = []
    for doc, score in results:
        if score < setting.score_threshold:
            continue
        records.append(Record(
            content=doc.page_content,
            score=round(score, 4),
            title=doc.metadata.get("source", "unknown"),
            metadata=doc.metadata,
        ))

    return RetrievalResponse(records=records)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("api:app", host="0.0.0.0", port=8000, reload=False)
