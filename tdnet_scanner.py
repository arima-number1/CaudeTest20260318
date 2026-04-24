#!/usr/bin/env python3
"""
TDnetリアルタイム監視エンジン v2.0 — アナリスト浅井
竹村トレードカンパニー

スケジュール: 毎時 00:01, 10:01, 20:01, 30:01, 40:01, 50:01 (10分おき)
出力先: outputs/asai/
"""

import os, sys, re, json, time, hashlib, logging
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from bs4 import BeautifulSoup
import anthropic
import yfinance as yf
import master_db

# ─── 設定 ────────────────────────────────────────────────────────────────
BASE_URL       = "https://www.release.tdnet.info/inbs/"
OUTPUT_DIR     = Path("outputs/asai")
SEEN_FILE      = OUTPUT_DIR / ".seen_ids.json"
ALERT_FILE     = OUTPUT_DIR / "weekend_alert.md"
PID_FILE       = OUTPUT_DIR / ".scanner_pid"
ALERT_THRESH   = 60   # |スコア|≥60 でアラート
JST            = ZoneInfo("Asia/Tokyo")

# スキャン起点分（毎時）
SCAN_MINUTES = [1, 11, 21, 31, 41, 51]

# 監視市場指数・コモディティ
MARKET_TICKERS = {
    "日経225":     "^N225",
    "TOPIX":       "1306.T",
    "S&P500":      "^GSPC",
    "NASDAQ":      "^IXIC",
    "ダウ平均":    "^DJI",
    "上海総合":    "000001.SS",
    "香港ハンセン": "^HSI",
    "USD/JPY":     "USDJPY=X",
    "EUR/JPY":     "EURJPY=X",
    "WTI原油":     "CL=F",
    "金(GOLD)":    "GC=F",
    "VIX":         "^VIX",
}

# ─── ロガー ──────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("浅井Scanner")


# ════════════════════════════════════════════════════════════════════════
#  マーケット概況取得
# ════════════════════════════════════════════════════════════════════════

def fetch_market_overview() -> dict:
    """主要指数・コモディティの直近データを取得"""
    overview = {}
    try:
        tickers = yf.Tickers(" ".join(MARKET_TICKERS.values()))
        for name, sym in MARKET_TICKERS.items():
            try:
                hist = tickers.tickers[sym].history(period="2d", interval="1d")
                if len(hist) >= 2:
                    prev  = hist["Close"].iloc[-2]
                    last  = hist["Close"].iloc[-1]
                    chg   = (last - prev) / prev * 100
                    overview[name] = {"price": round(last, 2), "change_pct": round(chg, 2)}
                elif len(hist) == 1:
                    last = hist["Close"].iloc[-1]
                    overview[name] = {"price": round(last, 2), "change_pct": 0.0}
            except Exception:
                overview[name] = {"price": None, "change_pct": None}
    except Exception as e:
        log.warning(f"市場データ取得エラー: {e}")
    return overview


def market_overview_md(overview: dict) -> str:
    """マーケット概況をMarkdownテーブルに整形"""
    if not overview:
        return "_市場データ取得不可_\n"

    lines = [
        "| 指数/銘柄 | 直近値 | 前日比 | トレンド |",
        "|-----------|--------|--------|---------|",
    ]
    for name, data in overview.items():
        price = data.get("price")
        chg   = data.get("change_pct")
        if price is None:
            lines.append(f"| {name} | — | — | — |")
            continue
        arrow = "▲" if chg and chg > 0 else ("▼" if chg and chg < 0 else "→")
        chg_s = f"{chg:+.2f}%" if chg is not None else "—"
        lines.append(f"| {name} | {price:,.2f} | {chg_s} | {arrow} |")
    return "\n".join(lines)


def market_sentiment(overview: dict) -> str:
    """マーケット全体のセンチメントをひと言で"""
    changes = [v["change_pct"] for v in overview.values() if v.get("change_pct") is not None]
    if not changes:
        return "データ不足"
    avg = sum(changes) / len(changes)
    vix = overview.get("VIX", {}).get("price")
    if vix and vix > 25:
        return f"⚠️ 高VIX({vix:.1f}) — リスクオフ局面"
    if avg > 0.5:
        return "🟢 リスクオン（全体的に上昇基調）"
    if avg < -0.5:
        return "🔴 リスクオフ（全体的に下落基調）"
    return "🟡 中立（方向感なし）"


# ════════════════════════════════════════════════════════════════════════
#  TDnetクローラー
# ════════════════════════════════════════════════════════════════════════

def fetch_page(date_str: str, page: int = 1) -> list[dict]:
    url = f"{BASE_URL}I_list_{page:03d}_{date_str}.html"
    try:
        resp = requests.get(url, timeout=20)
        resp.encoding = "utf-8"
    except Exception as e:
        log.warning(f"fetch failed {url}: {e}")
        return []
    soup = BeautifulSoup(resp.text, "lxml")
    table = soup.find("table", id="main-list-table")
    if not table:
        return []
    items = []
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 4:
            continue
        tag = tds[3].find("a")
        if not tag:
            continue
        code  = tds[1].get_text(strip=True)
        title = tag.get_text(strip=True)
        uid   = hashlib.md5(f"{date_str}{code}{title}".encode()).hexdigest()[:12]
        items.append({
            "id":    uid,
            "date":  date_str,
            "time":  tds[0].get_text(strip=True),
            "code":  code,
            "name":  tds[2].get_text(strip=True),
            "title": title,
            "place": tds[5].get_text(strip=True) if len(tds) > 5 else "",
        })
    return items


def fetch_recent(hours: int = 6) -> list[dict]:
    """直近N時間の開示を取得"""
    now = datetime.now(JST)
    dates = set()
    for h in range(hours + 1):
        dates.add((now - timedelta(hours=h)).strftime("%Y%m%d"))
    all_items = []
    for d in sorted(dates, reverse=True):
        for p in range(1, 6):
            pg = fetch_page(d, p)
            if not pg:
                break
            all_items.extend(pg)
            if len(pg) < 99:
                break
    return all_items


# ════════════════════════════════════════════════════════════════════════
#  浅井スコアリング（analyst_skills.md 準拠）
# ════════════════════════════════════════════════════════════════════════

def score_disclosures(items: list[dict], client: anthropic.Anthropic) -> list[dict]:
    if not items:
        return []
    BATCH = 30
    scored = []
    for i in range(0, len(items), BATCH):
        batch = items[i:i + BATCH]
        news_text = "\n".join(
            f"{j+1}. [{x['code']}]{x['name']} | {x['title']}"
            for j, x in enumerate(batch)
        )
        prompt = f"""あなたは竹村トレードカンパニーのアナリスト「浅井」です。
専門: ニュース・IR情報のリアルタイム解釈と市場センチメント定量評価。

社長からの厳命: 「買い材料」だけでなく「売り材料（ネガティブスコア）」を必ず特定・報告すること。
ポジティブな開示ばかりに偏った分析は不完全とみなす。

以下のTDnet（東証適時開示）を分析せよ。売り材料となる開示（下方修正/不祥事/ランサム/希薄化/損失/回収/調査委設置/支配株主異動/横領/サイバー攻撃/業績悪化/第三者割当/訴訟/行政処分等）には躊躇なくネガティブスコアを付けること。

【センチメントスコア基準】
+60〜+100: 強気買い   → 業績大幅上振れ/TOB/株式分割+増配/大型M&A
+20〜+59:  やや強気   → 増配/自社株買い/新事業/ポジティブIR
-19〜+19:  中立       → 訂正（軽微）/定期開示/想定内の月次
-20〜-59:  やや弱気   → 業績下方修正/製品回収/希薄化懸念/損失発生
-60〜-100: 強気売り   → 不祥事/粉飾疑い/重大損失/ランサムウェア/上場廃止リスク/調査委設置

【判断のポイント（売り材料）】
- 「特別調査委員会設置」→ -80〜-90（不正確定前でも最大リスク）
- 「ランサムウェア被害」→ -60〜-80（第2報以降は更に重く）
- 「業績予想の修正（下方）」→ -30〜-60（修正幅に応じて）
- 「新株予約権の大量行使」→ -40〜-60（希薄化リスク）
- 「有価証券評価損」→ -30〜-55
- 「情報流出・不正アクセス」→ -50〜-70
- 「支配株主の異動」→ -20〜-40（文脈次第）
- 「（訂正）」付き開示 → 元の開示内容によるが基本ネガティブ評価

【market_impact】
HIGH: 5%以上の株価変動が想定（売り材料でも HIGH は存在する）
MEDIUM: 1〜5%程度
LOW: 1%未満

必ずJSON配列のみを返してください（説明文・コードブロック不要）:
[
  {{
    "code": "コード",
    "name": "会社名",
    "sentiment_score": <-100〜100の整数。売り材料には必ずマイナス値を付けること>,
    "market_impact": {{"level": "HIGH/MEDIUM/LOW", "direction": "BULLISH/NEUTRAL/BEARISH", "time_horizon": "即日/1週間以内/1ヶ月以内"}},
    "news_summary": "要点1〜2文",
    "impact_reason": "市場インパクトの根拠（買い/売りの具体的根拠を明記）",
    "signal": "BUY/SELL/NEUTRAL"
  }},
  ...
]

開示リスト:
{news_text}
"""
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=4096,
                messages=[{"role": "user", "content": prompt}]
            )
            text = resp.content[0].text.strip()
            m = re.search(r'\[.*\]', text, re.DOTALL)
            if m:
                results = json.loads(m.group())
                for r, item in zip(results, batch):
                    item.update({
                        "score":         r.get("sentiment_score", 0),
                        "impact_level":  r.get("market_impact", {}).get("level", "LOW"),
                        "direction":     r.get("market_impact", {}).get("direction", "NEUTRAL"),
                        "time_horizon":  r.get("market_impact", {}).get("time_horizon", "1週間以内"),
                        "news_summary":  r.get("news_summary", ""),
                        "impact_reason": r.get("impact_reason", ""),
                        "signal":        r.get("signal", "NEUTRAL"),
                    })
            else:
                for item in batch:
                    item.update({"score": 0, "impact_level": "LOW", "direction": "NEUTRAL",
                                 "time_horizon": "—", "news_summary": "", "impact_reason": ""})
        except Exception as e:
            log.error(f"scoring error batch {i}: {e}")
            for item in batch:
                item.update({"score": 0, "impact_level": "LOW", "direction": "NEUTRAL",
                             "time_horizon": "—", "news_summary": "", "impact_reason": "スコアリング失敗"})
        scored.extend(batch)

    return sorted(scored, key=lambda x: abs(x.get("score", 0)), reverse=True)


# ════════════════════════════════════════════════════════════════════════
#  出力フォーマッター
# ════════════════════════════════════════════════════════════════════════

def impact_emoji(level: str, direction: str) -> str:
    if level == "HIGH":
        return "🔴" if direction == "BEARISH" else "🟢"
    if level == "MEDIUM":
        return "🟠" if direction == "BEARISH" else "🟡"
    return "⚪"


def write_scan_report(
    scan_no: int,
    scored: list[dict],
    overview: dict,
    new_count: int,
    total_count: int,
) -> Path:
    now     = datetime.now(JST)
    now_str = now.strftime("%Y%m%d_%H%M")
    path    = OUTPUT_DIR / f"live_scan_{now_str}.md"

    positives  = [x for x in scored if x.get("score", 0) > 0]
    negatives  = [x for x in scored if x.get("score", 0) < 0]
    high_items = [x for x in scored if x.get("impact_level") == "HIGH"]

    # 買い/売りリスト
    sell_high = sorted(
        [x for x in scored if x.get("score", 0) < -19],
        key=lambda x: x.get("score", 0)
    )
    buy_high = sorted(
        [x for x in scored if x.get("score", 0) > 19],
        key=lambda x: x.get("score", 0), reverse=True
    )

    lines = [
        f"# 浅井 TDnetスキャンレポート #{scan_no}",
        f"**実行日時:** {now.strftime('%Y年%m月%d日 %H:%M')} JST",
        f"**新着開示:** {new_count}件 / 直近6h総件数: {total_count}件",
        "",
        "---",
        "",
        "## 📊 マーケット概況",
        "",
        f"**センチメント:** {market_sentiment(overview)}",
        "",
        market_overview_md(overview),
        "",
        "---",
        "",
        "## 🔴 売り材料ランキング（ネガティブスコア）",
        "",
        "> 社長への売り推奨リスト。スコアが低いほど強い売りシグナル。",
        "",
        "| スコア | LV | コード | 銘柄名 | シグナル | 売り根拠 |",
        "|--------|-----|--------|--------|---------|---------|",
    ]

    if sell_high:
        for item in sell_high[:10]:
            s   = item.get("score", 0)
            lv  = item.get("impact_level", "LOW")
            sig = item.get("signal", "SELL")
            reason = item.get("impact_reason", "")[:45]
            lines.append(
                f"| **{s}** | {lv} | {item['code']} | {item['name']} | {sig} | {reason} |"
            )
    else:
        lines.append("| — | — | — | — | — | 新着の売り材料なし |")

    lines += [
        "",
        "---",
        "",
        "## 🟢 買い材料ランキング（ポジティブスコア）",
        "",
        "| スコア | LV | コード | 銘柄名 | シグナル | 買い根拠 |",
        "|--------|-----|--------|--------|---------|---------|",
    ]

    if buy_high:
        for item in buy_high[:10]:
            s   = item.get("score", 0)
            lv  = item.get("impact_level", "LOW")
            sig = item.get("signal", "BUY")
            reason = item.get("impact_reason", "")[:45]
            lines.append(
                f"| **+{s}** | {lv} | {item['code']} | {item['name']} | {sig} | {reason} |"
            )
    else:
        lines.append("| — | — | — | — | — | 新着の買い材料なし |")

    lines += [
        "",
        "---",
        "",
        "## 📋 全新着開示スコア一覧（買い・売り・中立 全件）",
        "",
        "| 時刻 | スコア | シグナル | LV | コード | 銘柄名 | 開示タイトル |",
        "|------|--------|---------|-----|--------|--------|------------|",
    ]

    for item in scored:
        s      = item.get("score", 0)
        sg     = "+" if s > 0 else ""
        signal = item.get("signal", "NEUTRAL")
        lv     = item.get("impact_level", "LOW")
        sig_icon = {"BUY": "🟢BUY", "SELL": "🔴SELL", "NEUTRAL": "⚪HOLD"}.get(signal, signal)
        lines.append(
            f"| {item.get('time','')} | **{sg}{s}** | {sig_icon} | {lv} | {item['code']} "
            f"| {item['name']} | {item['title'][:42]} |"
        )

    if not scored:
        lines.append("| — | — | — | — | — | — | 新着開示なし |")

    lines += [
        "",
        "---",
        "",
        "## 📢 全社共有（浅井 → 深井・不労・金田・豆田）",
        "",
        f"[報告] 浅井 が live_scan_{now_str}.md を出力しました。",
        f"スキャン#{scan_no} 完了 | 新着{new_count}件 | "
        f"HIGH: {len(high_items)}件 | "
        f"🟢買い推奨: {len(buy_high)}件 | 🔴売り推奨: {len(sell_high)}件 | "
        f"⚪中立: {len(scored)-len(buy_high)-len(sell_high)}件",
        "",
        "_本レポートは自動生成です。投資判断は社長の最終決裁に基づきます。_",
    ]

    path.write_text("\n".join(lines), encoding="utf-8")
    log.info(f"[出力] {path.name}  (HIGH:{len(high_items)} POS:{len(positives)} NEG:{len(negatives)})")
    return path


def append_alert(scan_no: int, scored: list[dict]):
    alerts = [x for x in scored if abs(x.get("score", 0)) >= ALERT_THRESH]
    if not alerts:
        return
    now_str = datetime.now(JST).strftime("%Y年%m月%d日 %H:%M")
    header_needed = not ALERT_FILE.exists()

    with open(ALERT_FILE, "a", encoding="utf-8") as f:
        if header_needed:
            f.write(
                f"# 週末アラート蓄積ログ — 浅井\n"
                f"**開始:** {now_str} | **閾値:** |スコア|≥{ALERT_THRESH}\n\n---\n"
            )
        f.write(f"\n## スキャン #{scan_no} — {now_str}  ({len(alerts)}件の重大開示)\n\n")
        sell_alerts = sorted([x for x in alerts if x.get("score", 0) < 0], key=lambda x: x["score"])
        buy_alerts  = sorted([x for x in alerts if x.get("score", 0) > 0], key=lambda x: -x["score"])
        for x in sell_alerts:
            s  = x.get("score", 0)
            lv = x.get("impact_level", "")
            f.write(
                f"- 🔴 **売り** [{lv}] **[{x['code']}] {x['name']}** "
                f"スコア:**{s}** | {x['title'][:50]}\n"
                f"  > {x.get('impact_reason','')}\n\n"
            )
        for x in buy_alerts:
            s  = x.get("score", 0)
            lv = x.get("impact_level", "")
            f.write(
                f"- 🟢 **買い** [{lv}] **[{x['code']}] {x['name']}** "
                f"スコア:**+{s}** | {x['title'][:50]}\n"
                f"  > {x.get('impact_reason','')}\n\n"
            )
    log.info(f"[アラート] {len(alerts)}件 → weekend_alert.md")


# ════════════════════════════════════════════════════════════════════════
#  スケジューラー
# ════════════════════════════════════════════════════════════════════════

def next_run_time() -> datetime:
    """次回スキャン時刻を計算（毎時 01,11,21,31,41,51分）"""
    now = datetime.now(JST)
    for m in SCAN_MINUTES:
        candidate = now.replace(minute=m, second=1, microsecond=0)
        if candidate > now + timedelta(seconds=5):
            return candidate
    # 次の時間の最初の分
    next_hour = (now + timedelta(hours=1)).replace(minute=SCAN_MINUTES[0], second=1, microsecond=0)
    return next_hour


def wait_until(target: datetime):
    now = datetime.now(JST)
    delta = (target - now).total_seconds()
    if delta > 0:
        log.info(f"次回スキャン: {target.strftime('%H:%M:%S')} (あと {int(delta//60)}分{int(delta%60)}秒)")
        time.sleep(delta)


# ════════════════════════════════════════════════════════════════════════
#  メインループ
# ════════════════════════════════════════════════════════════════════════

def load_seen() -> set:
    if SEEN_FILE.exists():
        return set(json.loads(SEEN_FILE.read_text()))
    return set()


def save_seen(seen: set):
    SEEN_FILE.write_text(json.dumps(list(seen)))


def run_scan(scan_no: int, client: anthropic.Anthropic, seen: set) -> set:
    log.info(f"═══ スキャン #{scan_no} 開始 ═══")

    # 市場概況取得
    log.info("市場概況取得中...")
    overview = fetch_market_overview()

    # TDnet取得
    log.info("TDnet取得中...")
    all_items  = fetch_recent(hours=6)
    new_items  = [x for x in all_items if x["id"] not in seen]
    log.info(f"新着: {len(new_items)}件 / 全件: {len(all_items)}件")

    # スコアリング
    scored = []
    if new_items:
        log.info(f"浅井スコアリング中... ({len(new_items)}件)")
        scored = score_disclosures(new_items, client)

    # レポート出力（新着がなくてもマーケット概況は出力）
    write_scan_report(scan_no, scored, overview, len(new_items), len(all_items))

    # アラート追記
    if scored:
        append_alert(scan_no, scored)

    # ── マスターDB更新 & インデックス再生成 ──────────────────────────
    if scored:
        added, total = master_db.ingest_and_rebuild(scored, source="TDnet")
        log.info(f"[DB] +{added}件追加 → 累計{total}件 / インデックス3本更新")

    # 既知ID更新
    for item in new_items:
        seen.add(item["id"])
    save_seen(seen)

    log.info(f"═══ スキャン #{scan_no} 完了 ═══\n")
    return seen


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        log.error("ANTHROPIC_API_KEY が設定されていません")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    # PIDファイル記録
    PID_FILE.write_text(str(os.getpid()))

    log.info("╔══════════════════════════════════════════╗")
    log.info("║  TDnet監視エンジン v2.0 — アナリスト浅井  ║")
    log.info("║  スケジュール: 毎時 01,11,21,31,41,51分   ║")
    log.info("╚══════════════════════════════════════════╝")
    log.info(f"出力先: {OUTPUT_DIR.resolve()}")

    seen     = load_seen()
    scan_no  = 0

    # 初回は即時実行
    scan_no += 1
    seen = run_scan(scan_no, client, seen)

    # 以降はスケジュール通りに実行
    while True:
        target = next_run_time()
        wait_until(target)
        scan_no += 1
        seen = run_scan(scan_no, client, seen)


if __name__ == "__main__":
    if "--once" in sys.argv:
        # GitHub Actions用：1回だけスキャンして終了
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            log.error("ANTHROPIC_API_KEY が設定されていません")
            sys.exit(1)
        client = anthropic.Anthropic(api_key=api_key)
        seen = load_seen()
        run_scan(1, client, seen)
    else:
        main()
