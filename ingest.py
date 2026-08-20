"""
ドキュメントを読み込み、ChromaDBにベクトル化して保存するスクリプト。
対応形式: PDF, DOCX, PPTX, XLSX
"""

import os
import glob
from pathlib import Path

from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_core.documents import Document

# PDF
from pypdf import PdfReader

# DOCX
from docx import Document as DocxDocument

# PPTX
from pptx import Presentation

# XLSX
import openpyxl


DOCS_DIR = Path(__file__).parent / "docs"
CHROMA_DIR = Path(__file__).parent / "chroma_db"
EMBED_MODEL = "intfloat/multilingual-e5-small"  # 多言語対応・軽量モデル


def load_pdf(path: str) -> str:
    reader = PdfReader(path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def load_docx(path: str) -> str:
    doc = DocxDocument(path)
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            row_text = "\t".join(cell.text.strip() for cell in row.cells if cell.text.strip())
            if row_text:
                parts.append(row_text)
    # テキストボックス（floating shape）
    for elem in doc.element.body.iter():
        if elem.tag.endswith('}txbxContent'):
            txbx_text = "".join(t.text for t in elem.iter() if t.tag.endswith('}t') and t.text)
            if txbx_text.strip():
                parts.append(txbx_text)
    return "\n".join(parts)


def load_pptx(path: str) -> str:
    prs = Presentation(path)
    texts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                for row in shape.table.rows:
                    row_text = "\t".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        texts.append(row_text)
            elif shape.has_chart:
                try:
                    chart = shape.chart
                    for plot in chart.plots:
                        cats = []
                        if hasattr(plot, "categories") and plot.categories:
                            cats = [str(c) if c is not None else "" for c in plot.categories]
                        for series in plot.series:
                            name = series.name or ""
                            vals = [str(v) if v is not None else "" for v in series.values]
                            if cats:
                                for cat, val in zip(cats, vals):
                                    if cat or val:
                                        texts.append(f"{name}\t{cat}\t{val}")
                            elif vals:
                                texts.append(f"{name}\t" + "\t".join(vals))
                except Exception:
                    pass
            elif hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text)
    return "\n".join(texts)


def load_xlsx(path: str) -> str:
    wb = openpyxl.load_workbook(path, data_only=True)
    texts = []
    for sheet in wb.worksheets:
        texts.append(f"[シート: {sheet.title}]")
        for row in sheet.iter_rows(values_only=True):
            row_text = "\t".join(str(c) for c in row if c is not None)
            if row_text.strip():
                texts.append(row_text)
    return "\n".join(texts)


LOADERS = {
    ".pdf": load_pdf,
    ".docx": load_docx,
    ".pptx": load_pptx,
    ".xlsx": load_xlsx,
}


def load_documents() -> list[Document]:
    documents = []
    for path in DOCS_DIR.iterdir():
        ext = path.suffix.lower()
        loader = LOADERS.get(ext)
        if loader is None:
            print(f"  スキップ（未対応形式）: {path.name}")
            continue
        print(f"  読み込み中: {path.name}")
        try:
            text = loader(str(path))
            if text.strip():
                documents.append(Document(
                    page_content=text,
                    metadata={"source": path.name, "file_type": ext.lstrip(".")}
                ))
        except Exception as e:
            print(f"  エラー ({path.name}): {e}")
    return documents


def main():
    print("=== ドキュメント読み込み開始 ===")
    raw_docs = load_documents()
    print(f"読み込み完了: {len(raw_docs)} ファイル\n")

    print("=== テキスト分割 ===")
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=500,
        chunk_overlap=50,
        separators=["\n\n", "\n", "。", "、", " ", ""],
    )
    chunks = splitter.split_documents(raw_docs)
    print(f"チャンク数: {len(chunks)}\n")

    print(f"=== 埋め込みモデル読み込み: {EMBED_MODEL} ===")
    embeddings = HuggingFaceEmbeddings(
        model_name=EMBED_MODEL,
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )

    print(f"=== ChromaDB に保存: {CHROMA_DIR} ===")
    vectorstore = Chroma.from_documents(
        documents=chunks,
        embedding=embeddings,
        persist_directory=str(CHROMA_DIR),
        collection_name="dify_rag",
    )
    print(f"保存完了: {vectorstore._collection.count()} チャンクを登録しました。")


if __name__ == "__main__":
    main()
