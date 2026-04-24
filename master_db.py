#!/usr/bin/env python3
"""
master_db.py — 浅井スコアリングDB管理モジュール
竹村トレードカンパニー

- master_db.json : 全スコア済み銘柄の永続ストア
- master_index_chronological.md : 新着順
- master_index_high_score.md    : スコア高い順
- master_index_low_score.md     : スコア低い順
"""

import json
import fcntl
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

OUTPUT_DIR = Path("outputs/asai")
DB_FILE    = OUTPUT_DIR / "master_db.json"
JST        = ZoneInfo("Asia/Tokyo")

IDX_CHRONO = OUTPUT_DIR / "master_index_chronological.md"
IDX_HIGH   = OUTPUT_DIR / "master_index_high_score.md"
IDX_LOW    = OUTPUT_DIR / "master_index_low_score.md"


# ════════════════════════════════════════════════════════════════════════
#  DB ロード / セーブ（ファイルロック付き）
# ════════════════════════════════════════════════════════════════════════

def load_db() -> dict:
    """DBを読み込む。形式: {uid: record}"""
    if not DB_FILE.exists():
        return {}
    with open(DB_FILE, "r", encoding="utf-8", errors="replace") as f:
        try:
            return json.load(f)
        except (json.JSONDecodeError, ValueError):
            return {}


def save_db(db: dict):
    """DBを書き込む（排他ロック）"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(DB_FILE, "w", encoding="utf-8") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        json.dump(db, f, ensure_ascii=False, indent=2)
        fcntl.flock(f, fcntl.LOCK_UN)


# ════════════════════════════════════════════════════════════════════════
#  レコード upsert
# ════════════════════════════════════════════════════════════════════════

def upsert_records(items: list[dict], source: str = "TDnet") -> int:
    """
    スコア済みアイテムリストをDBにupsert。
    同一IDが既存の場合はスコアが非0なら上書き。
    戻り値: 新規追加件数
    """
    db      = load_db()
    now_str = datetime.now(JST).strftime("%Y-%m-%d %H:%M")
    added   = 0

    for item in items:
        uid   = item.get("id") or _make_uid(item)
        score = item.get("score", 0)

        if uid in db and db[uid].get("score", 0) != 0 and score == 0:
            # 既存に有効スコアがある場合は上書きしない
            continue

        db[uid] = {
            "uid":          uid,
            "scan_time":    now_str,
            "date":         item.get("date", ""),
            "time":         item.get("time", ""),
            "code":         item.get("code", ""),
            "name":         item.get("name", ""),
            "title":        item.get("title", ""),
            "score":        score,
            "signal":       item.get("signal", "NEUTRAL"),
            "impact_level": item.get("impact_level", "LOW"),
            "direction":    item.get("direction", "NEUTRAL"),
            "time_horizon": item.get("time_horizon", "—"),
            "news_summary": item.get("news_summary", ""),
            "impact_reason":item.get("impact_reason", ""),
            "source":       source,
            "place":        item.get("place", ""),
        }
        added += 1

    save_db(db)
    return added


def _make_uid(item: dict) -> str:
    import hashlib
    key = f"{item.get('date','')}{item.get('code','')}{item.get('title','')}"
    return hashlib.md5(key.encode()).hexdigest()[:12]


# ════════════════════════════════════════════════════════════════════════
#  インデックス再生成
# ════════════════════════════════════════════════════════════════════════

def _sig_icon(signal: str, score: int) -> str:
    if signal == "BUY"  or score > 19:  return "🟢BUY"
    if signal == "SELL" or score < -19: return "🔴SELL"
    return "⚪HOLD"


def _lv_icon(level: str, score: int) -> str:
    icons = {"HIGH": "🔴HIGH" if score < 0 else "🟢HIGH",
             "MEDIUM": "🟠MED" if score < 0 else "🟡MED",
             "LOW": "⚪LOW"}
    return icons.get(level, "⚪LOW")


def _score_fmt(s: int) -> str:
    return f"+{s}" if s > 0 else str(s)


def _table_row(r: dict) -> str:
    s      = r.get("score", 0)
    sig    = _sig_icon(r.get("signal", "NEUTRAL"), s)
    lv     = _lv_icon(r.get("impact_level", "LOW"), s)
    reason = (r.get("impact_reason") or r.get("news_summary") or "")[:55]
    return (
        f"| {r.get('scan_time','')} "
        f"| **{_score_fmt(s)}** | {sig} | {lv} "
        f"| {r.get('code','')} | {r.get('name','')} "
        f"| {r.get('title','')[:38]} "
        f"| {reason} "
        f"| {r.get('source','TDnet')} |"
    )


def _header(title: str, db: dict, sort_key, reverse: bool) -> list[str]:
    now    = datetime.now(JST).strftime("%Y年%m月%d日 %H:%M")
    total  = len(db)
    sells  = sum(1 for r in db.values() if r.get("score", 0) < -19)
    buys   = sum(1 for r in db.values() if r.get("score", 0) > 19)
    return [
        f"# {title}",
        f"**最終更新:** {now} JST  |  **総件数:** {total}件  "
        f"|  🔴売り推奨: {sells}件  |  🟢買い推奨: {buys}件",
        "",
        "| スキャン日時 | スコア | シグナル | LV | コード | 銘柄名 | 開示タイトル | 根拠要約 | ソース |",
        "|------------|--------|---------|-----|--------|--------|------------|---------|--------|",
    ]


def rebuild_indexes():
    """3つのインデックスMDファイルを全件再生成"""
    db      = load_db()
    records = list(db.values())
    now     = datetime.now(JST).strftime("%Y年%m月%d日 %H:%M")
    total   = len(records)
    sells   = [r for r in records if r.get("score", 0) <= -20]
    buys    = [r for r in records if r.get("score", 0) >= 20]
    neutral = [r for r in records if -19 <= r.get("score", 0) <= 19]

    def footer():
        return [
            "",
            "---",
            "",
            "## 📢 全社共有（浅井 → 深井・不労・金田・豆田）",
            "",
            f"[報告] マスターインデックス更新完了 — {now} JST",
            f"総件数: {total}件 | 🔴売り: {len(sells)}件 | 🟢買い: {len(buys)}件 | ⚪中立: {len(neutral)}件",
            "",
            "_投資判断は社長の最終決裁に基づきます。_",
        ]

    COL = "| スキャン日時 | スコア | シグナル | LV | コード | 銘柄名 | 開示タイトル | 根拠要約 | ソース |"
    SEP = "|------------|--------|---------|-----|--------|--------|------------|---------|--------|"

    # ── 時系列順（新着→古い）────────────────────────────────────────────
    chrono = sorted(records, key=lambda r: r.get("scan_time", ""), reverse=True)
    lines  = [
        "# 浅井マスターインデックス — 時系列順（新着順）",
        f"**最終更新:** {now} JST  |  **総件数:** {total}件  "
        f"|  🔴売り: {len(sells)}件  |  🟢買い: {len(buys)}件",
        "",
        COL, SEP,
    ] + [_table_row(r) for r in chrono] + footer()
    IDX_CHRONO.write_text("\n".join(lines), encoding="utf-8")

    # ── スコア高い順（買い材料）──────────────────────────────────────────
    high = sorted(records, key=lambda r: r.get("score", 0), reverse=True)
    lines = [
        "# 浅井マスターインデックス — スコア高い順（買い材料ランキング）",
        f"**最終更新:** {now} JST  |  **総件数:** {total}件  "
        f"|  🟢買い推奨（≥+20）: {len(buys)}件",
        "",
        COL, SEP,
    ] + [_table_row(r) for r in high] + footer()
    IDX_HIGH.write_text("\n".join(lines), encoding="utf-8")

    # ── スコア低い順（売り材料）──────────────────────────────────────────
    low = sorted(records, key=lambda r: r.get("score", 0))
    lines = [
        "# 浅井マスターインデックス — スコア低い順（売り材料ランキング）",
        f"**最終更新:** {now} JST  |  **総件数:** {total}件  "
        f"|  🔴売り推奨（≤-20）: {len(sells)}件",
        "",
        COL, SEP,
    ] + [_table_row(r) for r in low] + footer()
    IDX_LOW.write_text("\n".join(lines), encoding="utf-8")

    return total, len(sells), len(buys)


# ════════════════════════════════════════════════════════════════════════
#  公開API
# ════════════════════════════════════════════════════════════════════════

def ingest_and_rebuild(items: list[dict], source: str = "TDnet") -> tuple[int, int]:
    """
    スコア済みリストをDBに追加し、3インデックスを再生成。
    戻り値: (追加件数, DB総件数)
    """
    added = upsert_records(items, source)
    total, sells, buys = rebuild_indexes()
    return added, total


if __name__ == "__main__":
    # 単体実行でインデックス再生成
    total, sells, buys = rebuild_indexes()
    print(f"インデックス再生成完了: {total}件 (売り:{sells} 買い:{buys})")
    print(f"  {IDX_CHRONO}")
    print(f"  {IDX_HIGH}")
    print(f"  {IDX_LOW}")
