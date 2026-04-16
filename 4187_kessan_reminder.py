#!/usr/bin/env python3
"""
4187 大阪有機化学工業 — 2Q決算前リマインド
2026/07/10 決算 → 7/8〜7/10 の朝 8:52 JST に実行
"""
import requests
import yfinance as yf
from datetime import datetime
from zoneinfo import ZoneInfo

JST = ZoneInfo("Asia/Tokyo")
now = datetime.now(JST)
today_str = now.strftime("%Y-%m-%d（%A）")

TEAMS_WEBHOOK = (
    "https://defaultbbe542ee9cf543d8beffd7181ba08c.38.environment.api.powerplatform.com"
    ":443/powerautomate/automations/direct/workflows/"
    "30db79c373a549adbc94df611f0c8ade/triggers/manual/paths/invoke"
    "?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0"
    "&sig=Ynn0E8zbujSXWZavhduLJjAwO7GXrtfZsN-aKhdDvxI"
)

KESSAN_DATE = datetime(2026, 7, 10, tzinfo=JST)
ENTRY_LOW   = 3800
ENTRY_HIGH  = 4100
TP          = 5200
STOP        = 3500


def fetch_price():
    try:
        t = yf.Ticker("4187.T")
        hist = t.history(period="30d")
        if not hist.empty:
            closes = hist["Close"]
            current = float(closes.iloc[-1])
            prev    = float(closes.iloc[-2])
            chg     = (current - prev) / prev * 100
            ma25    = float(closes.rolling(25).mean().iloc[-1])
            ma_diff = (current - ma25) / ma25 * 100
            delta = closes.diff()
            gain = delta.clip(lower=0).rolling(14).mean()
            loss = (-delta.clip(upper=0)).rolling(14).mean()
            rsi = float((100 - 100 / (1 + gain / loss)).iloc[-1])
            return current, chg, ma25, ma_diff, rsi
    except Exception:
        pass
    return None, None, None, None, None


def main():
    current, chg, ma25, ma_diff, rsi = fetch_price()
    days_left = (KESSAN_DATE - now).days

    price_str  = f"¥{current:,.0f}（{chg:+.2f}%）" if current else "取得失敗"
    upside_str = f"+{(TP - current) / current * 100:.1f}%" if current else "—"
    ma_str     = f"¥{ma25:,.0f}（乖離{ma_diff:+.1f}%）" if ma25 else "—"
    rsi_str    = f"{rsi:.0f}" if rsi else "—"

    # エントリーゾーン判定
    if current and current <= ENTRY_HIGH:
        signal = "🟢 エントリーゾーン内 → 押し目買いチャンス"
    elif current and current <= ENTRY_HIGH * 1.05:
        signal = "🟡 ゾーン近傍 → 引き続き押し目待ち"
    else:
        signal = "⚪ ゾーン上方 → 2Q発表待ち"

    msg = f"""【⏰ 4187 大阪有機化学工業 2Q決算リマインド】あと{days_left}日

📅 本日：{today_str}
🗓 2Q決算日：2026-07-10（木）

━━━━━━━━━━━━━━━━━━━━━━━━━━
{signal}

■ 現在株価：{price_str}
■ MA25：{ma_str}  RSI：{rsi_str}
■ TP：¥{TP:,}（現値比{upside_str}）  ストップ：¥{STOP:,}

━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 2Qの見方（深井DR 2026-04-16 より）
⚠️  1Qの棚卸再評価益（¥2〜3億）が消え、見た目▲25〜35%に見えやすい
✅ 本質（ArFモノマー需要・製品値上げ浸透）は堅固 → 市場の誤解が買い場

■ 投資テーゼ（3本柱）
① ArFモノマー世界シェア60〜70%寡占（JSR・信越化学・東京応化が顧客）
② EUV用モノマーが27.11期以降の成長ドライバー（AI/DC半導体微細化直結）
③ PEG<1.0・PER12.8x — OP成長+15%対比で明確に割安

■ 押し目エントリーゾーン：¥{ENTRY_LOW:,}〜¥{ENTRY_HIGH:,}
■ ストップライン：¥{STOP:,}（▲{(1 - STOP / (current or TP)) * 100:.1f}%）

_深井 竹村トレードカンパニー_"""

    try:
        r = requests.post(TEAMS_WEBHOOK, json={"text": msg}, timeout=15)
        print(f"Teams通知送信: HTTP {r.status_code} | 決算まで{days_left}日 | {price_str}")
    except Exception as e:
        print(f"通知エラー: {e}")


if __name__ == "__main__":
    main()
