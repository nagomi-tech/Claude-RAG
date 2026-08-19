"""
Vision-based image indexing for Q21-25 (embedded chart/diagram images).

Extracts images from:
  - 売上実績ダッシュボード_2024上半期.pdf       (pages 2-3: bar/pie charts)
  - 組織図_テックリードジャパン_2024年4月.pptx  (slide 2: org chart PNG)
  - 業務フロー図_カスタマーサクセス対応.docx    (inline image: flowchart PNG)

Sends each image to Claude Vision → extracts text → adds to ChromaDB.
"""

import base64
import io
import os
import time
from pathlib import Path

import anthropic
from dotenv import load_dotenv
from langchain_core.documents import Document
from langchain_chroma import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings

load_dotenv()

DOCS = Path(__file__).parent / "docs"
CHROMA_DIR = Path(__file__).parent / "chroma_db"
EMBED_MODEL = "intfloat/multilingual-e5-small"
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")


# ─── 画像抽出 ─────────────────────────────────────────────────────────────────

def pdf_pages_as_png(pdf_path: Path, page_numbers: list[int]) -> list[tuple[int, bytes]]:
    """Render specified PDF pages (1-indexed) as PNG bytes using pypdfium2."""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(pdf_path))
    results = []
    for page_no in page_numbers:
        page = pdf[page_no - 1]
        bitmap = page.render(scale=2.0)
        pil_image = bitmap.to_pil()
        buf = io.BytesIO()
        pil_image.save(buf, format="PNG")
        results.append((page_no, buf.getvalue()))
    return results


def pptx_picture_blobs(pptx_path: Path) -> list[tuple[str, bytes]]:
    """Extract Picture shape image bytes from each PPTX slide."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(str(pptx_path))
    images = []
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                images.append((f"スライド{i}", shape.image.blob))
    return images


def docx_inline_image_blobs(docx_path: Path) -> list[tuple[str, bytes]]:
    """Extract inline image bytes from DOCX."""
    from docx import Document

    doc = Document(str(docx_path))
    images = []
    for i, shape in enumerate(doc.inline_shapes, 1):
        try:
            rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            image_part = doc.part.related_parts[rId]
            images.append((f"図{i}", image_part.blob))
        except Exception:
            pass
    return images


# ─── Vision 説明文生成 ────────────────────────────────────────────────────────

def describe_image(client: anthropic.Anthropic, image_bytes: bytes,
                   doc_name: str, image_desc: str) -> str:
    """Call Claude Vision and return full text description of the image."""
    b64 = base64.standard_b64encode(image_bytes).decode()
    # Detect media type from bytes header
    media_type = "image/png" if image_bytes[:4] == b"\x89PNG" else "image/jpeg"

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": media_type, "data": b64},
                },
                {
                    "type": "text",
                    "text": (
                        f"これは「{doc_name}」の{image_desc}です。\n"
                        "この画像に含まれるすべての情報（数値、名前、項目名、構造、関係性など）を"
                        "漏れなく日本語のテキストとして書き起こしてください。\n"
                        "グラフの場合は各項目の値を、組織図の場合は階層構造と人名を、"
                        "フロー図の場合は各ステップと条件分岐を具体的に記述してください。"
                    ),
                },
            ],
        }],
    )
    return message.content[0].text


# ─── ChromaDB 追加 ────────────────────────────────────────────────────────────

def get_vectorstore() -> Chroma:
    embeddings = HuggingFaceEmbeddings(
        model_name=EMBED_MODEL,
        model_kwargs={"device": "cpu"},
        encode_kwargs={"normalize_embeddings": True},
    )
    return Chroma(
        persist_directory=str(CHROMA_DIR),
        embedding_function=embeddings,
        collection_name="dify_rag",
    )


# ─── メイン ───────────────────────────────────────────────────────────────────

def main():
    if not ANTHROPIC_API_KEY:
        raise SystemExit("Error: ANTHROPIC_API_KEY not set in .env")

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    vectorstore = get_vectorstore()

    tasks: list[tuple[str, str, bytes]] = []  # (source_title, image_desc, image_bytes)

    # ── PDF: ページ2（棒グラフ）・ページ3（円グラフ） ──────────────────────────
    pdf_path = DOCS / "売上実績ダッシュボード_2024上半期.pdf"
    doc_name = "売上実績ダッシュボード_2024上半期.pdf"
    for page_no, img_bytes in pdf_pages_as_png(pdf_path, [2, 3]):
        label = "月別売上高推移（棒グラフ）" if page_no == 2 else "事業部別売上比率（円グラフ）"
        tasks.append((doc_name, f"p.{page_no} {label}", img_bytes))
    print(f"  PDF: {len([t for t in tasks if '売上' in t[0]])}ページ抽出")

    # ── PPTX: 組織図スライド ───────────────────────────────────────────────────
    pptx_path = DOCS / "組織図_テックリードジャパン_2024年4月.pptx"
    doc_name_p = "組織図_テックリードジャパン_2024年4月.pptx"
    for slide_label, img_bytes in pptx_picture_blobs(pptx_path):
        tasks.append((doc_name_p, f"{slide_label} 組織図", img_bytes))
    print(f"  PPTX: {len([t for t in tasks if '組織図' in t[0]])}画像抽出")

    # ── DOCX: フロー図 ─────────────────────────────────────────────────────────
    docx_path = DOCS / "業務フロー図_カスタマーサクセス対応.docx"
    doc_name_d = "業務フロー図_カスタマーサクセス対応.docx"
    for fig_label, img_bytes in docx_inline_image_blobs(docx_path):
        tasks.append((doc_name_d, f"{fig_label} 問い合わせ対応フロー", img_bytes))
    print(f"  DOCX: {len([t for t in tasks if 'フロー' in t[0]])}画像抽出")

    print(f"\n合計 {len(tasks)} 件の画像をVisionで解析します...\n")

    docs_to_add: list[Document] = []
    for source, image_desc, img_bytes in tasks:
        print(f"  Vision解析中: {source} [{image_desc}]", end=" ", flush=True)
        text = describe_image(client, img_bytes, source, image_desc)
        docs_to_add.append(Document(
            page_content=text,
            metadata={"source": source, "file_type": "vision_extracted", "image_desc": image_desc},
        ))
        print(f"→ {len(text)}文字")
        time.sleep(1.0)

    vectorstore.add_documents(docs_to_add)
    print(f"\nChromaDBに {len(docs_to_add)} チャンク追加完了。")
    print("api.py を再起動してください。")


if __name__ == "__main__":
    main()
