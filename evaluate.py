"""
eval_qa.json の20問を使ってRAG APIの正答率を評価するスクリプト。

評価軸:
  1. ソース正答  : 期待するファイル名のチャンクがtop-k内に含まれるか
  2. キーワード正答: expected_answer 内の重要キーワードが取得チャンク内に含まれるか
  3. スコア分布  : 類似度スコアの統計情報
"""

import json
import re
import sys
import time
import urllib.request
import urllib.error
from pathlib import Path

API_URL    = "http://localhost:8000/retrieval"
API_KEY    = "dify-rag-secret"
TOP_K      = 5
THRESHOLD  = 0.2
QA_FILE    = Path(__file__).parent / "eval_qa.json"

# ─── キーワード抽出（数値・固有名詞・カタカナなど重要語を抜く）─────────────
def extract_keywords(text: str) -> list[str]:
    """expected_answer から評価に使うキーワードを抽出する"""
    keywords = []
    # 数値+単位 (例: 35.2億円, +12.4%, 850名, 5億円)
    keywords += re.findall(r'[\d,.]+(?:億円|万円|百万円|名|件|%|点|pt|ポイント|社)', text)
    # 固有名詞っぽいもの（カタカナ2文字以上）
    keywords += re.findall(r'[ァ-ヴー]{3,}', text)
    # 固有名詞（漢字＋カタカナ混じり、3文字以上）
    keywords += re.findall(r'[一-龯ァ-ヴA-Za-z]{3,}(?:Pro|AI|ERP|NPS|CSAT|IT)?', text)
    # アルファベット略語
    keywords += re.findall(r'\b[A-Z]{2,}\b', text)
    # 人名っぽいもの（苗字＋空白＋名前）
    keywords += re.findall(r'[一-龯]{1,3}[\s　][一-龯ぁ-ん]{1,3}', text)

    # 短すぎる・ストップワード除去
    stop = {"ます", "です", "した", "ある", "いる", "から", "まで", "また", "さらに",
            "および", "または", "として", "について", "それぞれ", "すべて"}
    keywords = [k.strip() for k in keywords if len(k.strip()) >= 2 and k.strip() not in stop]
    return list(dict.fromkeys(keywords))[:10]  # 重複除去・最大10語


# ─── API呼び出し ────────────────────────────────────────────────────────────
def call_retrieval(query: str) -> list[dict]:
    body = json.dumps({
        "query": query,
        "retrieval_setting": {"top_k": TOP_K, "score_threshold": THRESHOLD}
    }).encode()
    req = urllib.request.Request(
        API_URL,
        data=body,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read())["records"]
    except Exception as e:
        print(f"    [API ERROR] {e}")
        return []


# ─── 評価ロジック ────────────────────────────────────────────────────────────
def evaluate_source(records: list[dict], expected_source: str) -> tuple[bool, float]:
    """正しいソースファイルが取得できているか"""
    for rec in records:
        if expected_source in rec["title"]:
            return True, rec["score"]
    return False, 0.0


def evaluate_keywords(records: list[dict], expected_answer: str) -> tuple[float, list[str], list[str]]:
    """expected_answer のキーワードが取得チャンクに含まれるか"""
    keywords = extract_keywords(expected_answer)
    if not keywords:
        return 1.0, [], []

    all_content = "\n".join(r["content"] for r in records)
    hit, miss = [], []
    for kw in keywords:
        if kw in all_content:
            hit.append(kw)
        else:
            miss.append(kw)

    return len(hit) / len(keywords), hit, miss


# ─── メイン評価 ─────────────────────────────────────────────────────────────
def main():
    qa_list = json.loads(QA_FILE.read_text(encoding="utf-8"))
    answerable     = [q for q in qa_list if q.get("currently_answerable", True)]
    not_answerable = [q for q in qa_list if not q.get("currently_answerable", True)]

    print(f"評価開始: 全{len(qa_list)}問 / TOP_K={TOP_K} / threshold={THRESHOLD}")
    print(f"  現時点で取得可能: {len(answerable)}問  |  将来課題（未対応要素）: {len(not_answerable)}問\n")
    print("=" * 80)

    results = []

    for qa in qa_list:
        qid           = qa["id"]
        question      = qa["question"]
        expected_src  = qa["source"]
        expected_ans  = qa["expected_answer"]
        category      = qa["category"]
        difficulty    = qa["difficulty"]
        elem_type     = qa.get("element_type", "本文テキスト")
        is_answerable = qa.get("currently_answerable", True)

        label = "〇" if is_answerable else "✕未対応"
        print(f"Q{qid:02d} [{difficulty}][{label}] {question[:55]}...")
        if not is_answerable:
            print(f"     未対応要素: {elem_type}")
            print(f"     改善方針 : {qa.get('improvement_needed', '─')[:70]}")

        records = call_retrieval(question)
        time.sleep(0.3)

        if not records:
            print(f"     → 結果なし\n")
            results.append({"id": qid, "source_hit": False, "kw_ratio": 0.0,
                            "top_score": 0.0, "difficulty": difficulty, "category": category,
                            "element_type": elem_type, "currently_answerable": is_answerable})
            continue

        source_hit, src_score = evaluate_source(records, expected_src)
        kw_ratio, hit_kws, miss_kws = evaluate_keywords(records, expected_ans)

        top_score = records[0]["score"] if records else 0.0
        top_src   = records[0]["title"] if records else "─"

        src_mark = "✓" if source_hit else "✗"
        kw_mark  = "✓" if kw_ratio >= 0.5 else ("△" if kw_ratio > 0 else "✗")

        print(f"     ソース: {src_mark} (期待:{expected_src[:30]})")
        print(f"     取得1位: {top_src[:35]} (score={top_score:.4f})")
        print(f"     KW一致: {kw_mark} {kw_ratio*100:.0f}%  hit={hit_kws[:3]}  miss={miss_kws[:3]}")
        print()

        results.append({
            "id": qid, "category": category, "difficulty": difficulty,
            "element_type": elem_type, "currently_answerable": is_answerable,
            "source_hit": source_hit, "src_score": src_score,
            "top_score": top_score, "kw_ratio": kw_ratio,
            "hit_kws": hit_kws, "miss_kws": miss_kws,
        })

    # ─── 集計 ────────────────────────────────────────────────────────────
    n      = len(results)
    curr   = [r for r in results if r["currently_answerable"]]
    future = [r for r in results if not r["currently_answerable"]]

    def stats(subset, label):
        if not subset: return
        sh  = sum(1 for r in subset if r["source_hit"])
        k50 = sum(1 for r in subset if r["kw_ratio"] >= 0.5)
        kp  = sum(1 for r in subset if r["kw_ratio"] == 1.0)
        sc  = [r["top_score"] for r in subset if r["top_score"] > 0]
        avg = sum(sc)/len(sc) if sc else 0
        nn  = len(subset)
        print(f"  {label} ({nn}問):")
        print(f"    ソース正答率      : {sh}/{nn} = {sh/nn*100:.1f}%")
        print(f"    KW正答率(50%以上) : {k50}/{nn}  = {k50/nn*100:.1f}%")
        print(f"    KW完全正答率      : {kp}/{nn}  = {kp/nn*100:.1f}%")
        print(f"    平均類似度スコア  : {avg:.4f}")

    src_hits   = sum(1 for r in results if r["source_hit"])
    kw_50      = sum(1 for r in results if r["kw_ratio"] >= 0.5)
    kw_perfect = sum(1 for r in results if r["kw_ratio"] == 1.0)
    scores     = [r["top_score"] for r in results if r["top_score"] > 0]
    avg_score  = sum(scores) / len(scores) if scores else 0

    print("=" * 80)
    print("【総合評価結果】\n")
    print(f"  評価問数（全体）  : {n} 問")
    print(f"  ソース正答率      : {src_hits}/{n} = {src_hits/n*100:.1f}%")
    print(f"  KW正答率(50%以上) : {kw_50}/{n}  = {kw_50/n*100:.1f}%")
    print(f"  KW完全正答率      : {kw_perfect}/{n}  = {kw_perfect/n*100:.1f}%")
    print(f"  平均類似度スコア  : {avg_score:.4f}")
    print()

    # 現在可能 / 将来課題 に分けて集計
    stats(curr,   "現時点で取得可能な問題")
    print()
    stats(future, "将来課題（未対応要素を含む問題）")
    print()

    # 要素タイプ別
    elem_types = list(dict.fromkeys(r["element_type"] for r in results))
    print("  要素タイプ別 KW正答率(50%以上):")
    for et in elem_types:
        sub = [r for r in results if r["element_type"] == et]
        kh  = sum(1 for r in sub if r["kw_ratio"] >= 0.5)
        ans = "✓取得可" if sub[0]["currently_answerable"] else "✕未対応"
        print(f"    {ans} {et:<22}: {kh}/{len(sub)} ({kh/len(sub)*100:.0f}%)")

    print()

    # 難易度別
    for diff in ["易", "中", "難"]:
        sub = [r for r in results if r["difficulty"] == diff]
        sh  = sum(1 for r in sub if r["source_hit"])
        kh  = sum(1 for r in sub if r["kw_ratio"] >= 0.5)
        print(f"  [{diff}] {len(sub)}問: ソース正答 {sh}/{len(sub)} ({sh/len(sub)*100:.0f}%)  KW正答 {kh}/{len(sub)} ({kh/len(sub)*100:.0f}%)")

    print()

    # 失敗した問題の一覧
    failed_curr = [r for r in curr if not r["source_hit"]]
    if failed_curr:
        print(f"  ✗ 【現在対応可能なのに】ソース取得失敗 ({len(failed_curr)}問):")
        for r in failed_curr:
            q = next(q for q in qa_list if q["id"] == r["id"])
            print(f"    Q{r['id']:02d} [{r['difficulty']}] {q['question'][:55]}...")
    else:
        print("  ✓ 現在対応可能な問題はすべて正しいソースドキュメントを取得できました")

    # 将来課題のまとめ
    print()
    print(f"  ── 将来課題（未対応要素）{len(future)}問の改善ロードマップ ──")
    for q in qa_list:
        if not q.get("currently_answerable", True):
            print(f"    Q{q['id']:02d} [{q['element_type']}]")
            print(f"         {q.get('improvement_needed', '─')[:72]}")

    print("\n" + "=" * 80)

    # JSON保存
    out = Path(__file__).parent / "eval_result.json"
    out.write_text(json.dumps({"summary": {
        "total": n, "source_hit": src_hits, "source_hit_rate": round(src_hits/n, 4),
        "kw_hit_50": kw_50, "kw_hit_rate_50": round(kw_50/n, 4),
        "kw_perfect": kw_perfect, "kw_perfect_rate": round(kw_perfect/n, 4),
        "avg_similarity_score": round(avg_score, 4),
    }, "details": results}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"詳細結果を保存しました: {out}")


if __name__ == "__main__":
    main()
