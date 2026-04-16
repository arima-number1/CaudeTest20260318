#!/usr/bin/env python3
"""
2884 ヨシムラ・フード — 中国ホタテ解禁ウォッチ
毎朝 8:52 JST に実行 → 解禁シグナル検出時は緊急BUY通知、未検出は監視継続通知
"""
import requests
import yfinance as yf
from datetime import datetime
from zoneinfo import ZoneInfo

JST = ZoneInfo("Asia/Tokyo")
today_str = datetime.now(JST).strftime("%Y-%m-%d（%A）")

TEAMS_WEBHOOK = (
    "https://defaultbbe542ee9cf543d8beffd7181ba08c.38.environment.api.powerplatform.com"
    ":443/powerautomate/automations/direct/workflows/"
    "30db79c373a549adbc94df611f0c8ade/triggers/manual/paths/invoke"
    "?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0"
    "&sig=Ynn0E8zbujSXWZavhduLJjAwO7GXrtfZsN-aKhdDvxI"
)

# ホタテ解禁キーワード（日本語・英語）
KAIKIN_KEYWORDS = [
    "ホタテ", "hotate", "scallop",
    "解禁", "輸入再開", "禁止解除", "輸入規制解除",
    "China lifts", "China resumes", "ban lifted", "import ban",
    "中国 輸入", "水産物 解禁",
]
BUY_TRIGGER_KEYWORDS = ["解禁", "輸入再開", "禁止解除", "輸入規制解除",
                         "lifts ban", "resumes imports", "ban lifted"]

ENTRY_ZONE_LOW  = 700
ENTRY_ZONE_HIGH = 800
TP              = 1200
STOP            = 600


def fetch_news_headlines():
    """Yahoo Finance の 2884.T ニュース取得"""
    headlines = []
    try:
        ticker = yf.Ticker("2884.T")
        news = ticker.news or []
        for item in news[:10]:
            title = item.get("title", "")
            if title:
                headlines.append(title)
    except Exception:
        pass
    return headlines


def check_kaikin(headlines):
    """解禁シグナル判定（ホタテ文脈 + 解禁ワード）"""
    combined = " ".join(headlines).lower()
    has_hotate = any(k.lower() in combined for k in ["ホタテ", "hotate", "scallop"])
    has_kaikin = any(k.lower() in combined for k in BUY_TRIGGER_KEYWORDS)
    return has_hotate and has_kaikin


def fetch_stock_price():
    try:
        ticker = yf.Ticker("2884.T")
        hist = ticker.history(period="5d")
        if not hist.empty:
            current = float(hist["Close"].iloc[-1])
            prev    = float(hist["Close"].iloc[-2])
            chg     = (current - prev) / prev * 100
            return current, chg
    except Exception:
        pass
    return None, None


def main():
    headlines = fetch_news_headlines()
    kaikin_detected = check_kaikin(headlines)
    current, chg = fetch_stock_price()

    price_str = f"¥{current:,.0f}（{chg:+.2f}%）" if current else "取得失敗"
    upside    = f"+{(TP - current) / current * 100:.1f}%" if current else "—"

    if kaikin_detected:
        # ========== 緊急 BUY アラート ==========
        headline_excerpt = "\n".join(f"  ・{h}" for h in headlines[:5])
        msg = f"""🚨【緊急 BUY シグナル】2884 ヨシムラ・フード — 中国ホタテ解禁検出！

📅 {today_str}

━━━━━━━━━━━━━━━━━━━━━━━━━━
🟢 中国ホタテ輸入解禁の可能性あり → 即時買い検討

■ 現在株価：{price_str}
■ エントリーゾーン：¥{ENTRY_ZONE_LOW:,}〜¥{ENTRY_ZONE_HIGH:,}
■ TP：¥{TP:,}（現値比{upside}）
■ ストップ：¥{STOP:,}

📰 検出ヘッドライン：
{headline_excerpt}

━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 投資テーゼ（深井 DR 2026-04-16）
① ヨシムラモデル（中小食品M&A）の構造テーゼは変わらず
② 中国ホタテ解禁 → 売上・粗利の直接回復
③ 有利子負債348億のキャッシュフロー改善加速
④ 現値¥750は決算ショック後の過剰売り水準

⚠️ 注意：ヘッドライン確認・IR発表を必ず目視確認の上エントリー判断

_深井 竹村トレードカンパニー_"""

    else:
        # ========== 日次監視報告 ==========
        hl_str = "\n".join(f"  ・{h}" for h in headlines[:3]) if headlines else "  （ニュースなし）"
        msg = f"""【🦪 2884 ヨシムラ・フード ホタテ解禁ウォッチ】{today_str}

🔴 中国ホタテ解禁シグナル：未検出（監視継続）

■ 現在株価：{price_str}
■ 待機水準：¥{ENTRY_ZONE_LOW:,}〜¥{ENTRY_ZONE_HIGH:,}
■ TP：¥{TP:,}（現値比{upside}）  ストップ：¥{STOP:,}

📰 直近ニュース：
{hl_str}

━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 買いトリガー条件（いずれか）
① 中国 水産物輸入規制解除・ホタテ輸入再開の発表
② 会社IR：中国向け輸出再開コメント
③ 株価¥650割れ → ナンピン検討（ストップ確認後）

_深井 竹村トレードカンパニー_"""

    try:
        r = requests.post(TEAMS_WEBHOOK, json={"text": msg}, timeout=15)
        status = "🚨緊急BUYアラート" if kaikin_detected else "監視継続"
        print(f"Teams通知送信: HTTP {r.status_code} | {status} | {price_str}")
    except Exception as e:
        print(f"通知エラー: {e}")


if __name__ == "__main__":
    main()
