#!/usr/bin/env python3
"""
bogdbank_source.py — Богд банкны улирлын санхүүгийн тайланг bogdbank.com-оос авна.

Сайт тайлангийн ЖАГСААЛТ өгдөггүй — он/улирал сонгоход AJAX POST хийж тухайн үеийн
тайланг HTML хүснэгтээр буцаадаг:
    POST https://www.bogdbank.com/report/getReportContent
         year=2026&session=3&report_category_id=2&report_type=1
    session            = улирал (1–4)
    report_category_id = 2  (https://www.bogdbank.com/report/2 — санхүүгийн тайлан)
    report_type        = 1  Монголбанкны стандартын дагуу (2 = IFRS)
    Дүн: МЯНГАН төгрөг (pipeline workbook-ийн өмнөх утгатай харьцуулж сая ₮ болгоно)

Иймд workbook-ийн сүүлийн улирлаас хойших улирлуудыг (өнөөдрийг хүртэл дууссан)
нэг нэгээр нь АСУУЖ, хүснэгттэй хариу ирсэн хамгийн сүүлийнхийг авна. Ижил улирлыг
ч дахин татаж hash-ийг харьцуулна (банк тайлангаа засварлавал pipeline дахин бичнэ).
bg_ci_session күүкийг эхлээд /report/2 хуудсыг ачаалж автоматаар авна.

    python3 bogdbank_source.py --probe 2026 3            # нэг улирлыг асууж задлаад хэвлэнэ
    python3 bogdbank_source.py --probe 2026 3 --save resp.html
    python3 bogdbank_source.py --from-html resp.html     # сүлжээгүй, хадгалсан хариугаар
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402

BANK = "Богд банк"
TICKER = "BOGD"
SITE = "https://www.bogdbank.com"
PAGE = f"{SITE}/report/2"
ENDPOINT = f"{SITE}/report/getReportContent"
REPORT_CATEGORY_ID = 2
REPORT_TYPE = 1            # 1 = Монголбанкны стандарт, 2 = IFRS
HEADERS = {
    "Accept": "*/*",
    "Accept-Language": "mn,en-US;q=0.9,en;q=0.8",
    "Origin": SITE,
    "Referer": PAGE,
    "X-Requested-With": "XMLHttpRequest",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36"),
}

_SESSION: requests.Session | None = None


def session() -> requests.Session:
    global _SESSION
    if _SESSION is None:
        s = requests.Session()
        s.headers.update(HEADERS)
        try:
            s.get(PAGE, timeout=30)          # bg_ci_session күүки авах
        except Exception:  # noqa: BLE001
            pass
        _SESSION = s
    return _SESSION


def fetch_report(year: int, quarter: int, report_type: int = REPORT_TYPE) -> str:
    r = session().post(ENDPOINT, data={"year": year, "session": quarter,
                                       "report_category_id": REPORT_CATEGORY_ID,
                                       "report_type": report_type},
                       headers={"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"},
                       timeout=60)
    r.raise_for_status()
    return r.text


def _html_payload(text: str) -> str:
    """Хариу HTML эсвэл {"html": "..."} / {"data": "..."} JSON байж болно."""
    t = text.strip()
    if t[:1] in "{[":
        try:
            js = json.loads(t)
        except ValueError:
            return text
        parts = []

        def walk(o):
            if isinstance(o, str) and ("<t" in o or "<div" in o):
                parts.append(o)
            elif isinstance(o, dict):
                for v in o.values():
                    walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)
        walk(js)
        return "\n".join(parts) or text
    return text


def parse_html(text: str) -> dict[str, float]:
    """Хүснэгтийн мөр бүр: эхний кирилл (>3 тэмдэгт) нүд = шошго, түүнээс баруун
    талын ХАМГИЙН СҮҮЛИЙН тоо = тайлант үе (Excel задлагчтай ижил дүрэм)."""
    soup = BeautifulSoup(_html_payload(text), "html.parser")
    raw: dict[str, float] = {}
    seen = set()
    for tr in soup.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        li = next((i for i, c in enumerate(cells) if len(c) > 3 and bot.CYRILLIC_RE.search(c)), None)
        if li is None:
            continue
        nums = [bot.parse_numeric_cell(c.replace(" ", " ").replace(" ", "")) for c in cells[li + 1:]]
        nums = [n for n in nums if n is not None]
        if not nums:
            continue
        label = re.sub(r"\s+", " ", cells[li]).strip(" .:")
        n = bot.normalize_row_label(label)
        if n in bot.SECTION_HEADER_LABELS or n in seen:
            continue
        seen.add(n)
        raw[label] = nums[-1]
    bot.aggregate_npl_categories(raw)
    return raw


def parse_file(path: Path) -> dict[str, float]:
    """.html (сайтын хариу) → parse_html; .xlsx/.xls/.pdf (inbox-д гараар) → ерөнхий задлагч."""
    path = Path(path)
    if path.suffix.lower() in (".html", ".htm"):
        return parse_html(path.read_text(encoding="utf-8", errors="replace"))
    import capitron_source
    return capitron_source.parse_file(path)


def adapt_to_sheet(raw: dict[str, float]) -> dict[str, float]:
    return raw


def _next(p):
    return (p[0], p[1] + 1) if p[1] < 4 else (p[0] + 1, 1)


def _last_closed_quarter(today: date | None = None):
    t = today or date.today()
    q = (t.month - 1) // 3          # одоогийн улирлаас өмнөх
    return (t.year, q) if q else (t.year - 1, 4)


MIN_LABELS = 8


def candidates_for(min_period, reports_dir: Path, today: date | None = None):
    """workbook-ийн (min_period) улирлаас эхлээд дууссан улирал бүрийг асууна.
    Буцаах: (хүснэгттэй хариу ирсэн хамгийн сүүлийн улирал, нэр дэвшигчид, тайлбар)."""
    last = _last_closed_quarter(today)
    probes, p = [], min_period
    while p <= last:
        probes.append(p)
        p = _next(p)
    probes = probes or [min_period]
    best = None
    for p in reversed(probes):                      # шинээс хуучин руу
        try:
            text = fetch_report(*p)
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠ Богд банк {p[0]}Q{p[1]}: {e}")
            continue
        raw = parse_html(text)
        if len(raw) >= MIN_LABELS:
            best = (p, text, raw)
            break
        print(f"  · Богд банк {p[0]}Q{p[1]}: тайлан хараахан алга ({len(raw)} шошго)")
    if best is None:
        return None, [], "тайлан олдсонгүй"
    p, text, raw = best
    content = text.encode("utf-8")
    sha = hashlib.sha256(content).hexdigest()
    d = reports_dir / f"{p[0]}Q{p[1]}"
    d.mkdir(parents=True, exist_ok=True)
    path = d / f"{TICKER}_{sha[:10]}.html"
    if not path.exists():
        path.write_bytes(content)
    print(f"  · {path.name}: {len(raw)} шошго")
    cand = {"url": f"{ENDPOINT}?year={p[0]}&session={p[1]}&report_type={REPORT_TYPE}",
            "path": path, "sha": sha, "raw": raw, "dates": {p}, "audited": False}
    return p, [cand], f"{p[0]} оны {p[1]}-р улирал (Монголбанкны стандарт)"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--probe", nargs=2, type=int, metavar=("YEAR", "QUARTER"))
    ap.add_argument("--ifrs", action="store_true", help="report_type=2 (IFRS)")
    ap.add_argument("--save", default=None)
    ap.add_argument("--from-html", default=None)
    a = ap.parse_args()
    if a.from_html:
        text = Path(a.from_html).read_text(encoding="utf-8")
    elif a.probe:
        text = fetch_report(a.probe[0], a.probe[1], 2 if a.ifrs else 1)
        if a.save:
            Path(a.save).write_text(text, encoding="utf-8")
            print(f"✓ хариу -> {a.save}")
    else:
        ap.error("--probe YEAR QUARTER эсвэл --from-html FILE")
    raw = parse_html(text)
    print(f"{len(raw)} шошго (мянган ₮):")
    for k, v in raw.items():
        print(f"  {k[:70]:70} {v:>20,.1f}")


if __name__ == "__main__":
    main()
