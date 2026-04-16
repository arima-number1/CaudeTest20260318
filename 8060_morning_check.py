#!/usr/bin/env python3
"""
8060 キヤノンMJ 決算前押し目チェック
毎朝 8:47 JST に実行 → Teams通知
決算日: 2026-04-22
"""
import requests
import yfinance as yf
import pandas as pd
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

KESSAN_DATE = datetime(2026, 4, 22, tzinfo=JST)
ENTRY_REF   = 3503   # DR想定価格（4/16時点）
DIPS_LOW    = 3400   # 押し目ゾーン下限
DIPS_HIGH   = 3450   # 押し目ゾーン上限
TP          = 4150   # Deep Research TP
STOP        = 3000   # ストップライン

def calc_rsi(closes, period=14):
    delta = closes.diff()
    gain = delta.clip(lower=0).rolling(period).mean()
    loss = (-delta.clip(upper=0)).rolling(period).mean()
    rs = gain / loss
    return 100 - (100 / (1 + rs))

def main():
    # 株価取得
    ticker = yf.Ticker("8060.T")
    hist = ticker.history(period="60d")
    if hist.empty:
        print("株価取得失敗")
        return

    closes = hist["Close"]
    current = float(closes.iloc[-1])
    prev    = float(closes.iloc[-2])
    chg_pct = (current - prev) / prev * 100

    # MA25 / RSI14
    ma25 = float(closes.rolling(25).mean().iloc[-1])
    ma_diff_pct = (current - ma25) / ma25 * 100
    rsi = float(calc_rsi(closes).iloc[-1])

    # 判定
    days_left = (KESSAN_DATE - datetime.now(JST)).days
    tp_upside = (TP - current) / current * 100

    if current <= DIPS_HIGH:
        action = "🟢 押し目ゾーン到達 → 買いチャンス"
        action_detail = f"¥{current:,.0f} は押し目ゾーン（¥{DIPS_LOW:,}〜¥{DIPS_HIGH:,}）内。即時エントリー検討。"
    elif current <= ENTRY_REF:
        action = "🟡 DR参照価格以下 → 検討圏"
        action_detail = f"¥{current:,.0f} はDR想定（¥{ENTRY_REF:,}）以下。MA乖離{ma_diff_pct:+.1f}%・RSI{rsi:.0f}で過熱感なし。"
    elif current <= ENTRY_REF * 1.03:
        action = "🟡 小幅上昇中 → 引き続き押し目待ち"
        action_detail = f"¥{current:,.0f} はDR想定を{(current-ENTRY_REF)/ENTRY_REF*100:+.1f}%上回り。押し目ゾーン¥{DIPS_HIGH:,}まで▲{(current-DIPS_HIGH)/current*100:.1f}%。"
    elif rsi > 75:
        action = "🔴 過熱 → 押し目待ち継続"
        action_detail = f"RSI{rsi:.0f}で過熱域。急追いは禁物。決算まで残り{days_left}日。"
    else:
        action = "⚪ 上昇継続 → 引き続き押し目待ち"
        action_detail = f"¥{current:,.0f}（+{chg_pct:+.1f}%）で上昇中。押し目ゾーン¥{DIPS_HIGH:,}まで▲{(current-DIPS_HIGH)/current*100:.1f}%。"

    msg = f"""【⏰ 8060 キヤノンMJ 押し目リマインド】決算まであと{days_left}日

📅 本日：{today_str}
🗓 決算日：2026-04-22（水）引け後

━━━━━━━━━━━━━━━━━━━━━━━━━━
{action}

■ 現在株価：¥{current:,.0f}（前日比{chg_pct:+.2f}%）
■ MA25：¥{ma25:,.0f}（乖離{ma_diff_pct:+.1f}%）
■ RSI14：{rsi:.0f}
■ TP：¥{TP:,}（現値比+{tp_upside:.1f}%）  ストップ：¥{STOP:,}

💡 {action_detail}

━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 投資テーゼ（深井DR 2026-04-16）
① IT販社→SIerへのPER再評価（精密14x→IT22x）
② 中堅企業（GB事業部）開拓で新規ストック収益拡大
③ OPM 8.6%→9.8%（FY25→30E）× EPS CAGR +6.7%
④ 純現預金¥1,395億（D/E=0x）が下値支持

■ 押し目ゾーン：¥{DIPS_LOW:,}〜¥{DIPS_HIGH:,}（MA25前後）
■ ストップライン：¥{STOP:,}（▲14.4%）

_深井 竹村トレードカンパニー_"""

    try:
        r = requests.post(TEAMS_WEBHOOK, json={"text": msg}, timeout=15)
        print(f"Teams通知送信: HTTP {r.status_code}")
        print(f"8060現在値: ¥{current:,.0f} | RSI:{rsi:.0f} | MA乖離:{ma_diff_pct:+.1f}% | 決算まで{days_left}日")
    except Exception as e:
        print(f"通知エラー: {e}")

if __name__ == "__main__":
    main()
