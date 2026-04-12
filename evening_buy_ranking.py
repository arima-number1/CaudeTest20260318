#!/usr/bin/env python3
"""
evening_buy_ranking.py — 深井 夕方BUY優先順位ランキング
竹村トレードカンパニー

毎営業日夕方（17:00 JST）に自動実行。
deep_research_index.csv のBUY/STRONG BUY銘柄を対象に：
  1. Yahoo Finance でリアルタイム株価取得
  2. TP乖離率を再計算
  3. 複合スコアで優先順位付け
  4. Teamsへ通知
"""

import csv, os, re, requests, json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from typing import Optional

JST        = ZoneInfo("Asia/Tokyo")
BASE_DIR   = Path(__file__).parent
INDEX_CSV  = BASE_DIR / "outputs/fukai/deep_research_index.csv"
MASTER_DIR = BASE_DIR / "inputs/raw_data"
TEAMS_WEBHOOK = (
    "https://defaultbbe542ee9cf543d8beffd7181ba08c.38.environment.api.powerplatform.com"
    ":443/powerautomate/automations/direct/workflows/"
    "30db79c373a549adbc94df611f0c8ade/triggers/manual/paths/invoke"
    "?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0"
    "&sig=Ynn0E8zbujSXWZavhduLJjAwO7GXrtfZsN-aKhdDvxI"
)

# ── 銘柄マスターCSVから次回決算日を自動取得 ────────────────
def load_kessan_dates() -> dict:
    """
    inputs/raw_data/銘柄マスター*.csv の最新ファイルから
    {コード: "YYYY-MM-DD"} の辞書を返す。
    コード列フォーマット: "6918 JT Equity" → "6918"
    日付列フォーマット: "2026/05/14" → "2026-05-14"
    """
    import glob
    files = sorted(glob.glob(str(MASTER_DIR / "銘柄マスター*.csv")))
    if not files:
        return {}
    latest = files[-1]
    dates = {}
    try:
        with open(latest, encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                # BOMつきキー対応
                code_raw = row.get("\ufeff銘柄コード", row.get("銘柄コード", "")).strip()
                m = re.match(r"^([A-Z0-9]{3,5})\s", code_raw)
                if not m:
                    continue
                code = m.group(1)
                raw_date = row.get("次回決算発表日", "").strip()
                if raw_date:
                    # "2026/05/14" → "2026-05-14"
                    dates[code] = raw_date.replace("/", "-")
    except Exception as e:
        print(f"  [銘柄マスター] 読み込み失敗: {e}")
    return dates

# ── テーマ分類 ─────────────────────────────────────────────
THEMES = {
    "有事・ホルムズ": ["9104","9101","9119","1605","1662","9107"],
    "AI半導体":       ["6361","4047","277A","6806","6161","4629","285A"],
    "防衛":           ["7012","6507","6269","6016"],
    "内需・小売":     ["9983","3994","3048","7085","2670"],
    "金融":           ["8316"],
    "バリュー低PBR":  ["5233","9324","9302","9672","3597","6287","5471"],
    "素材・化学":     ["4189","4631","5401","5713","5711","5706","5714"],
    "医薬・ヘルスケア":["4901","4534"],
}

def get_theme(code: str) -> str:
    for theme, codes in THEMES.items():
        if code in codes:
            return theme
    return "その他"

def fetch_price(code: str) -> Optional[float]:
    is_us = bool(re.match(r'^[A-Z]{2,5}$', code))
    symbol = code if is_us else f"{code}.T"
    headers = {"User-Agent": "Mozilla/5.0"}
    for base in ["https://query1.finance.yahoo.com", "https://query2.finance.yahoo.com"]:
        try:
            url = f"{base}/v8/finance/chart/{symbol}"
            r = requests.get(url, headers=headers, timeout=8)
            if r.status_code == 200:
                return float(r.json()["chart"]["result"][0]["meta"]["regularMarketPrice"])
        except Exception:
            pass
    return None

def days_to_kessan(code: str, kessan_dates: dict) -> Optional[int]:
    d = kessan_dates.get(code)
    if not d:
        return None
    try:
        target = datetime.strptime(d, "%Y-%m-%d").date()
        today  = datetime.now(JST).date()
        diff   = (target - today).days
        return diff if diff >= 0 else None
    except Exception:
        return None

def to_float(s) -> Optional[float]:
    try:
        return float(str(s).replace(",", "").replace("+", "").replace("%", "").strip())
    except Exception:
        return None

def composite_score(upside: Optional[float], days: Optional[int], theme: str, judgment: str) -> float:
    """
    優先スコア算出（高いほど優先）
    = 乖離率 + カタリスト近接ボーナス + テーマボーナス + STRONG BUYボーナス
    """
    score = 0.0

    # 乖離率（最大40点）
    if upside is not None:
        score += min(upside * 0.4, 40.0)

    # 決算近接ボーナス（最大30点：7日以内=30, 14日以内=20, 30日以内=10）
    if days is not None:
        if days <= 7:
            score += 30.0
        elif days <= 14:
            score += 20.0
        elif days <= 30:
            score += 10.0
        elif days <= 60:
            score += 5.0

    # テーマボーナス（有事継続局面）
    HIGH_PRIORITY_THEMES = {"有事・ホルムズ", "AI半導体", "防衛"}
    if theme in HIGH_PRIORITY_THEMES:
        score += 15.0

    # STRONG BUYボーナス
    if "STRONG" in judgment.upper():
        score += 10.0

    return round(score, 1)

def load_buy_stocks():
    rows = []
    if not INDEX_CSV.exists():
        return rows

    seen = set()  # 重複除去（同一コードは最新行を使用）
    all_rows = []
    with open(INDEX_CSV, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            all_rows.append(row)

    # 最新行を優先（後ろから処理してseen管理）
    # ※ コードが一度でも出現したら（HOLD/SELL含む）seenに登録して古いBUYエントリーを排除
    for row in reversed(all_rows):
        judgment = row.get("judgment", "").strip().upper()
        code     = row.get("code", "").strip()
        if not code:
            continue
        if code in seen:
            continue
        # 最新判定が何であれ、このコードは処理済みとしてマーク
        seen.add(code)
        # BUY / STRONG BUY のみ追加（HOLD / SELL / WATCH / 依頼中 を除外）
        if "BUY" not in judgment:
            continue
        if any(x in judgment for x in ["HOLD", "SELL", "WATCH"]):
            continue
        rows.append(row)

    return list(reversed(rows))  # 日付順に戻す

def run():
    jst_now = datetime.now(JST)
    today   = jst_now.strftime("%Y-%m-%d")
    print(f"\n{'═'*100}")
    print(f"  📊 深井 夕方BUY優先順位ランキング  {today} {jst_now.strftime('%H:%M')} JST")
    print(f"  ※ バンス米副大統領 米イラン協議決裂（2026-04-12）→ 有事継続・ホルムズ封鎖継続を基本シナリオ")
    print(f"{'═'*100}\n")

    stocks = load_buy_stocks()
    if not stocks:
        print("  [ERROR] BUY銘柄が見つかりません。")
        return

    # 銘柄マスターから決算日を自動取得
    kessan_dates = load_kessan_dates()
    master_file  = sorted(__import__("glob").glob(str(MASTER_DIR / "銘柄マスター*.csv")))
    master_label = Path(master_file[-1]).name if master_file else "不明"
    print(f"  決算日ソース: {master_label} ({len(kessan_dates)}銘柄分)")

    # リアルタイム株価取得 & スコア算出
    ranked = []
    print(f"  株価取得中... ({len(stocks)}銘柄)")
    for row in stocks:
        code     = row["code"].strip()
        name     = row["name"].strip()
        judgment = row["judgment"].strip()
        tp_str   = row.get("target_price", "").strip()
        sp_str   = row.get("stock_price", "").strip()
        note     = row.get("notes", "").strip()[:60]

        tp = to_float(tp_str)

        # 株価：リアルタイム取得（失敗時はCSV記録値を使用）
        live_price = fetch_price(code)
        sp = live_price if live_price else to_float(sp_str)

        # 乖離率
        upside = None
        if tp and sp and sp > 0:
            upside = (tp - sp) / sp * 100

        theme   = get_theme(code)
        days    = days_to_kessan(code, kessan_dates)
        score   = composite_score(upside, days, theme, judgment)

        ranked.append({
            "code":      code,
            "name":      name,
            "judgment":  judgment,
            "tp":        tp,
            "sp":        sp,
            "live":      live_price is not None,
            "upside":    upside,
            "theme":     theme,
            "days":      days,
            "score":     score,
            "note":      note,
        })

    # スコア降順ソート
    ranked.sort(key=lambda x: -x["score"])

    # ── 出力 ────────────────────────────────────────────────
    header = (
        f"  {'順':>3}  {'コード':<6}  {'銘柄名':<22}  {'判定':<11}  "
        f"{'現在値':>8}  {'TP':>8}  {'乖離':>7}  {'決算':>6}  {'テーマ':<14}  スコア"
    )
    sep = "  " + "-"*97
    print(header)
    print(sep)

    teams_lines = [f"📊 **深井 夕方BUYランキング {today}**\n"]
    for i, s in enumerate(ranked, 1):
        sp_s  = f"¥{s['sp']:>8,.0f}" if s["sp"] else "  取得失敗"
        tp_s  = f"¥{s['tp']:>8,.0f}" if s["tp"] else "        -"
        up_s  = f"{s['upside']:>+6.1f}%" if s["upside"] is not None else "      -"
        day_s = f"{s['days']:>3}日" if s["days"] is not None else "   -"
        live_mark = "●" if s["live"] else "○"

        line = (
            f"  {i:>3}  {s['code']:<6}  {s['name']:<22}  {s['judgment']:<11}  "
            f"{live_mark}{sp_s}  {tp_s}  {up_s}  {day_s}  {s['theme']:<14}  {s['score']:>5.1f}"
        )
        print(line)

        # Teams用（上位15位まで）
        if i <= 15:
            up_t = f"{s['upside']:+.1f}%" if s["upside"] is not None else "-"
            day_t = f"{s['days']}日後" if s["days"] is not None else "-"
            teams_lines.append(
                f"{i}. **{s['code']} {s['name']}** [{s['judgment']}]  "
                f"現値:{sp_s.strip()} / TP:{tp_s.strip()} / 乖離:{up_t} / 決算:{day_t} / {s['theme']}"
            )

    # ── Tier表示 ────────────────────────────────────────────
    print(f"\n{'─'*100}")
    print("  🔴 Tier S（スコア55+）今日〜今週 仕込み急ぐ")
    tier_s = [s for s in ranked if s["score"] >= 55]
    for s in tier_s:
        up_s = f"{s['upside']:+.1f}%" if s["upside"] is not None else "-"
        sp_s = f"¥{s['sp']:,.0f}" if s["sp"] else "要確認"
        print(f"    ★ {s['code']} {s['name']:<20}  現値:{sp_s}  乖離:{up_s}  {s['note'][:55]}")

    print("\n  🟠 Tier A（スコア35〜54）来週〜分割買い")
    tier_a = [s for s in ranked if 35 <= s["score"] < 55]
    for s in tier_a:
        up_s = f"{s['upside']:+.1f}%" if s["upside"] is not None else "-"
        sp_s = f"¥{s['sp']:,.0f}" if s["sp"] else "要確認"
        print(f"    ◆ {s['code']} {s['name']:<20}  現値:{sp_s}  乖離:{up_s}  {s['note'][:55]}")

    print("\n  🟡 Tier B（スコア35未満）中長期・バリュー蓄積")
    tier_b = [s for s in ranked if s["score"] < 35]
    for s in tier_b:
        up_s = f"{s['upside']:+.1f}%" if s["upside"] is not None else "-"
        sp_s = f"¥{s['sp']:,.0f}" if s["sp"] else "要確認"
        print(f"    ◇ {s['code']} {s['name']:<20}  現値:{sp_s}  乖離:{up_s}  {s['note'][:55]}")

    print(f"\n{'═'*100}")
    print(f"  ● リアルタイム株価  ○ CSV記録値（取得失敗）")
    print(f"  スコア算出: 乖離率(max40) + 決算近接(max30) + テーマ(15) + STRONG BUY(10)")
    print(f"{'═'*100}\n")

    # ── Teams通知 ─────────────────────────────────────────
    teams_body = "\n".join(teams_lines)
    try:
        payload = {"text": teams_body}
        r = requests.post(TEAMS_WEBHOOK, json=payload, timeout=10)
        print(f"  [Teams] 通知送信: HTTP {r.status_code}")
    except Exception as e:
        print(f"  [Teams] 通知失敗: {e}")

if __name__ == "__main__":
    run()
