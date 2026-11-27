#!/usr/bin/env python3
"""
fetch_prices.py — 5 банкны (SBM, XAC, TDB, GLMT, KHAN) арилжааны түүхийг
old.mse.mn-ээс татаж, data/price_data.json-той нэгтгэнэ.

Компанийн ID (open.mse.mn/securities/{ID}) = ҮЦ код өөрөө:
    564 = Төрийн банк (SBM)
    568 = ХасБанк (XAC)
    567 = ХХБ (TDB)
    562 = Голомт (GLMT)
    563 = Хаан банк (KHAN)
(2026 оны 9-р сард Claude-аар шууд баталгаажуулсан: 562, 563, 567, 568 амжилттай
татагдсан бөгөөд яг ҮЦ кодтойгоо тэнцүү болохыг харсан. 564 (SBM) мөн адил
хэв маягтай тул итгэлтэй, гэхдээ шууд fetch-ээр батлаагүй тул ажиллуулсны дараа
`--check` горимоор эхний удаад заавал шалгана уу.)

Ажиллуулах:
    pip install -r requirements.txt
    python3 scripts/fetch_prices.py            # бүрэн түүхийг татаж нэгтгэнэ
    python3 scripts/fetch_prices.py --check     # татахгүй, зөвхөн одоогийн байдлыг хэвлэнэ
"""
import re
import sys
import json
import time
from pathlib import Path
from datetime import datetime

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://open.mse.mn/securities/"
REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_FILE = REPO_ROOT / "data" / "price_data.json"

BANKS = {
    "Төрийн банк": {"ticker": "SBM", "id": 564},
    "ХасБанк":     {"ticker": "XAC", "id": 568},
    "ХХБ":         {"ticker": "TDB", "id": 567},
    "Голомт":      {"ticker": "GLMT", "id": 562},
    "Хаан банк":   {"ticker": "KHAN", "id": 563},
}

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "mn,en;q=0.5",
    "Referer": f"{BASE_URL}/mn",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# № Ханш Арилжигдсан Дээд Доод Нээлт Хаалт Ширхэг Үнийн дүн Огноо
ROW_PATTERN = re.compile(
    r"(\d+)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+"
    r"([\d,]+(?:\.\d+)?)\s+([\d,]+(?:\.\d+)?)\s+([\d,]+)\s+"
    r"[\d,]*(?:\.\d+)?\s+(\d{4}-\d{2}-\d{2})"
)


def fetch_company_history(company_id: int) -> list[dict]:
    url = f"{BASE_URL}/mn/company/{company_id}"
    resp = SESSION.get(url, timeout=30)
    resp.raise_for_status()
    text = resp.text

    records, seen = [], set()
    for m in ROW_PATTERN.finditer(text):
        try:
            high = float(m.group(2).replace(",", ""))
            low = float(m.group(3).replace(",", ""))
            open_ = float(m.group(4).replace(",", ""))
            close = float(m.group(5).replace(",", ""))
            vol = float(m.group(6).replace(",", ""))
            date = m.group(7)
        except ValueError:
            continue
        if open_ == 0 and close == 0:  # block trade — candlestick-ыг гажуудуулна
            continue
        if date in seen:
            continue
        seen.add(date)
        records.append({"date": date, "open": open_, "high": high,
                         "low": low, "close": close, "volume": vol})

    records.sort(key=lambda r: r["date"])
    return records


def load_existing() -> dict:
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    return {}


def to_columnar(records: list[dict]) -> dict:
    return {
        "date": [r["date"] for r in records],
        "open": [round(r["open"], 2) for r in records],
        "high": [round(r["high"], 2) for r in records],
        "low": [round(r["low"], 2) for r in records],
        "close": [round(r["close"], 2) for r in records],
        "volume": [int(r["volume"]) for r in records],
    }


def from_columnar(col: dict) -> list[dict]:
    n = len(col.get("date", []))
    return [{"date": col["date"][i], "open": col["open"][i], "high": col["high"][i],
              "low": col["low"][i], "close": col["close"][i], "volume": col["volume"][i]}
             for i in range(n)]


def merge(old: list[dict], new: list[dict]):
    existing = {r["date"] for r in old}
    added = [r for r in new if r["date"] not in existing]
    combined = sorted(old + added, key=lambda r: r["date"])
    return combined, len(added)


def run_check():
    price = load_existing()
    for bank, info in BANKS.items():
        recs = price.get(bank, {"date": []})
        n = len(recs["date"])
        last_date = recs["date"][-1] if n else "—"
        last_close = recs["close"][-1] if n else "—"
        print(f"  {info['ticker']:5} ({bank:12}): {n:4} өдөр | сүүлийн: {last_date}  хаалт={last_close}")


def run_fetch():
    price = load_existing()
    total_added = 0
    any_error = False
    for bank, info in BANKS.items():
        print(f"→ {bank} ({info['ticker']}, id={info['id']}) татаж байна...")
        try:
            history = fetch_company_history(info["id"])
        except Exception as e:
            print(f"  ⚠ Алдаа: {e}")
            any_error = True
            continue
        if not history:
            print("  ⚠ Хоосон хариу ирлээ — хуудасны бүтэц өөрчлөгдсөн байж болзошгүй")
            any_error = True
            continue
        old = from_columnar(price.get(bank, {"date": [], "open": [], "high": [], "low": [], "close": [], "volume": []}))
        combined, added = merge(old, history)
        price[bank] = to_columnar(combined)
        total_added += added
        last = combined[-1]
        print(f"  {len(old)} → {len(combined)} өдөр (+{added}) | сүүлийн: {last['date']}  хаалт={last['close']}")
        time.sleep(1.5)  # эелдэг байдлаар татах, сервэрт даралт үзүүлэхгүй

    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(price, ensure_ascii=False), encoding="utf-8")
    print(f"\n✓ {DATA_FILE} бичигдлээ (нийт +{total_added} шинэ өдөр)")

    # GitHub Actions-д алдаа гарсныг мэдэгдэх (workflow "continue but flag" хийж чадна)
    if any_error:
        sys.exit(2)


if __name__ == "__main__":
    if "--check" in sys.argv:
        run_check()
    else:
        run_fetch()
