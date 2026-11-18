#!/usr/bin/env python3
"""
ownership_source.py — ҮЦТХТ-ийн "Үнэт цаасны төвлөрлийн судалгаа"-ны нэгтгэл Excel-ээс
(github.com/consultingapartment/Huvitsaanii-Tuvlurul → data/uctht_concentration.xlsx)
МХБ-ийн 5 банкны хувьцаа эзэмшлийн мэдээллийг авна.

Тэр repo өөрийн workflow-оор сар бүр шинэчлэгддэг. Энэ модуль банкны тайлан гарах бүрт:
    • файлыг GitHub-аас татаж data/ownership/-д кэшлэнэ (татаж чадахгүй бол кэшийг ашиглана)
    • бүх "Тайлан YY-MM-DD" хуудсыг БАГАНЫН ГАРЧГААР нь уншина (баганын дараалал
      өөрчлөгдсөн ч ажиллана), банкийг ҮЦ КОДООР нь (= МХБ-ийн компанийн дугаар) олно
    • банк бүрт: гаргасан хувьцаа, төрийн мэдлийн %, 5%-иас дээш эзэмшигчдийн % ба тоо,
      чөлөөт эргэлт (5%-иас бага) %, эзэмшигчдийн тоо, 3/12 сарын өсөлт, эзэмшлийн
      бүтцийн ӨӨРЧЛӨЛТИЙН үйл явдлууд (давамгай хувь, төрийн хувь, хувьцааны тоо)

    python3 ownership_source.py                    # татаж, 5 банкны хураангуйг хэвлэнэ
    python3 ownership_source.py --file local.xlsx  # локал файлаас
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_URL = ("https://raw.githubusercontent.com/consultingapartment/Huvitsaanii-Tuvlurul/"
               "main/data/uctht_concentration.xlsx")
DEFAULT_CACHE = Path("data/ownership/uctht_concentration.xlsx")

# ҮЦ код = МХБ-ийн компанийн дугаар (bank_financials_bot.BANKS[...]["id"])
BANK_CODES = {"Төрийн банк": 564, "ХасБанк": 568, "ХХБ": 567, "Голомт": 562, "Хаан банк": 563}

SHEET_RE = re.compile(r"^Тайлан\s+(\d{2})-(\d{2})-(\d{2})$")
# Эзэмшлийн бүтцийн өөрчлөлтийг "үйл явдал" гэж тооцох босго
DOM_CHANGE_PP = 0.005          # давамгай хувь ≥0.5 нэгж хувиар
HOLDER_SURGE = 0.10            # эзэмшигчдийн тоо сард ≥10%


# ── Татах / кэш ────────────────────────────────────────────────────────────
def fetch(url: str = DEFAULT_URL, cache: Path = DEFAULT_CACHE, timeout: int = 60) -> tuple[Path, str]:
    """(файлын зам, төлөв). Төлөв: 'шинэ' | 'өөрчлөлтгүй' | 'кэш (татаж чадсангүй: …)'."""
    import requests
    cache = Path(cache)
    cache.parent.mkdir(parents=True, exist_ok=True)
    try:
        r = requests.get(url, timeout=timeout)
        r.raise_for_status()
        if r.content[:2] != b"PK":
            raise ValueError("xlsx биш хариу")
        old = hashlib.sha256(cache.read_bytes()).hexdigest() if cache.exists() else None
        new = hashlib.sha256(r.content).hexdigest()
        if old != new:
            cache.write_bytes(r.content)
            return cache, "шинэ"
        return cache, "өөрчлөлтгүй"
    except Exception as e:  # noqa: BLE001
        if cache.exists():
            return cache, f"кэш (татаж чадсангүй: {e})"
        raise RuntimeError(f"Эзэмшлийн файл татагдсангүй, кэш ч алга: {e}") from e


# ── Задлах ────────────────────────────────────────────────────────────────
def _n(s) -> str:
    return re.sub(r"[\s\-]+", "", str(s or "")).lower()


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace(" ", ""))
    except ValueError:
        return None


def _columns(ws) -> tuple[int, dict[str, int]]:
    """Гарчгийн 2 мөрөөс баганын зураглал. Буцаах: (өгөгдөл эхлэх мөр, {түлхүүр: багана})."""
    top = next((r for r in range(1, 15) if any(_n(ws.cell(r, c).value) == "үцкод"
                                               for c in range(1, ws.max_column + 1))), None)
    if top is None:
        raise ValueError(f"{ws.title}: 'ҮЦ код' гарчиг олдсонгүй")
    cols, group = {}, ""
    for c in range(1, ws.max_column + 1):
        t, s = _n(ws.cell(top, c).value), _n(ws.cell(top + 1, c).value)
        if t:
            group = t
        if t == "үцкод":
            cols["code"] = c
        elif "үнэтцааснынэр" in t:
            cols["name"] = c
        elif t.startswith("нийтгаргасан"):
            cols["total"] = c
        elif "төрийн" in group:
            if "хувь" in s:
                cols["state_pct"] = c
            elif "ширхэг" in s:
                cols["state_shares"] = c
        elif "5аасих" in group or "5%аасих" in group:
            key = "pct" if "хувь" in s else "holders" if "эзэмшигч" in s else "shares" if "ширхэг" in s else None
            if key:
                cols[f"dom_{key}"] = c
        elif "5аасбага" in group or "5%аасбага" in group:
            key = "pct" if "хувь" in s else "holders" if "эзэмшигч" in s else "shares" if "ширхэг" in s else None
            if key:
                cols[f"ff_{key}"] = c
    need = {"code", "total", "dom_pct", "ff_pct", "ff_holders"}
    if not need <= cols.keys():
        raise ValueError(f"{ws.title}: багана дутуу {sorted(need - cols.keys())}")
    return top + 2, cols


def load_snapshots(path: Path, codes: dict[str, int] = BANK_CODES) -> dict[str, list[dict]]:
    """{банк: [{date, total, state_pct, dom_pct, dom_holders, dom_shares, ff_pct, ff_shares, ff_holders}, …]}
    dom_* = 5%-иас ДЭЭШ эзэмшигчид, ff_* = 5%-иас ДООШ (чөлөөт эргэлт)
    огноогоор эрэмбэлсэн."""
    wb = load_workbook(path, data_only=True, read_only=False)
    by_code = {v: k for k, v in codes.items()}
    out: dict[str, list[dict]] = {b: [] for b in codes}
    for ws in wb.worksheets:
        m = SHEET_RE.match(ws.title.strip())
        if not m:
            continue
        d = date(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
        start, cols = _columns(ws)
        for r in range(start, ws.max_row + 1):
            code = _num(ws.cell(r, cols["code"]).value)
            if code is None or int(code) not in by_code:
                continue
            g = {k: _num(ws.cell(r, c).value) for k, c in cols.items() if k not in ("code", "name")}
            for k_ in ("dom_holders", "ff_holders", "dom_shares", "ff_shares", "total"):
                if g.get(k_) is not None:
                    g[k_] = int(round(g[k_]))      # тоо ширхэг — бүхэл
            rec = {"date": d.isoformat(), "name": ws.cell(r, cols.get("name", 1)).value,
                   "total": g.get("total"), "state_pct": g.get("state_pct") or 0.0,
                   "dom_pct": g.get("dom_pct"), "dom_holders": g.get("dom_holders"),
                   "dom_shares": g.get("dom_shares"), "ff_pct": g.get("ff_pct"),
                   "ff_shares": g.get("ff_shares"), "ff_holders": g.get("ff_holders")}
            out[by_code[int(code)]].append(rec)
    for b in out:
        out[b].sort(key=lambda x: x["date"])
    return out


def _at_or_before(snaps: list[dict], iso: str) -> dict | None:
    best = None
    for s in snaps:
        if s["date"] <= iso:
            best = s
    return best


def _months_back(iso: str, months: int) -> str:
    y, m, d = map(int, iso.split("-"))
    m -= months
    while m <= 0:
        m += 12
        y -= 1
    return date(y, m, min(d, 28 if m == 2 else 30)).isoformat()


def summarize(snaps: dict[str, list[dict]], quarter_end_iso: str | None = None) -> dict[str, dict]:
    """Банк бүрийн сүүлийн байдал — 5%-иас ДЭЭШ ба ДООШ эзэмшигчдийг ТУСАД НЬ:
        over5  : {holders, pct, shares, holders_chg_3m/12m (тоо), pct_chg_3m/12m (н.х.)}
        under5 : {holders, pct, shares, holders_3m/12m (өсөлт %), pct_chg_3m/12m, avg_shares}
    + өөрчлөлтийн үйл явдлууд. Хуучин кодтой нийцүүлэх dom_*/ff_* түлхүүрүүд хэвээр."""
    out = {}
    for bank, ss in snaps.items():
        if not ss:
            continue
        cur = ss[-1]
        q = _at_or_before(ss, quarter_end_iso) if quarter_end_iso else None
        m3 = _at_or_before(ss, _months_back(cur["date"], 3))
        m12 = _at_or_before(ss, _months_back(cur["date"], 12))
        m3 = None if m3 is cur else m3
        m12 = None if m12 is cur else m12

        def diff(key, ref):
            return (cur[key] - ref[key]) if ref and None not in (cur.get(key), ref.get(key)) else None

        def grow(key, ref):
            return (cur[key] / ref[key] - 1) if ref and cur.get(key) and ref.get(key) else None

        events = []
        for a, b in zip(ss, ss[1:]):
            if b["date"] < _months_back(cur["date"], 12):
                continue
            if a["dom_pct"] is not None and b["dom_pct"] is not None and abs(b["dom_pct"] - a["dom_pct"]) >= DOM_CHANGE_PP:
                events.append({"date": b["date"], "type": "over5_pct", "group": "over5",
                               "text": f"5%-иас ДЭЭШ эзэмшигчдийн эзлэх хувь {a['dom_pct']:.1%} → {b['dom_pct']:.1%} "
                                       f"(5%-иас доош: {a['ff_pct']:.1%} → {b['ff_pct']:.1%})"})
            if (a.get("dom_holders") or 0) != (b.get("dom_holders") or 0):
                events.append({"date": b["date"], "type": "over5_count", "group": "over5",
                               "text": f"5%-иас ДЭЭШ эзэмшигчийн тоо {a['dom_holders']:.0f} → {b['dom_holders']:.0f}"})
            if abs((b["state_pct"] or 0) - (a["state_pct"] or 0)) >= DOM_CHANGE_PP:
                events.append({"date": b["date"], "type": "state", "group": "state",
                               "text": f"Төрийн мэдлийн хувь {a['state_pct']:.1%} → {b['state_pct']:.1%}"})
            if a["total"] and b["total"] and a["total"] != b["total"]:
                events.append({"date": b["date"], "type": "shares", "group": "total",
                               "text": f"Гаргасан хувьцаа {a['total']:,.0f} → {b['total']:,.0f}"})
            if a.get("ff_holders") and b.get("ff_holders") and b["ff_holders"] / a["ff_holders"] - 1 >= HOLDER_SURGE:
                events.append({"date": b["date"], "type": "under5_surge", "group": "under5",
                               "text": f"5%-иас ДООШ эзэмшигчдийн тоо {a['ff_holders']:,.0f} → {b['ff_holders']:,.0f} "
                                       f"({b['ff_holders'] / a['ff_holders'] - 1:+.0%})"})
        over5 = {"holders": cur["dom_holders"], "pct": cur["dom_pct"], "shares": cur["dom_shares"],
                 "holders_chg_3m": diff("dom_holders", m3), "holders_chg_12m": diff("dom_holders", m12),
                 "pct_chg_3m": diff("dom_pct", m3), "pct_chg_12m": diff("dom_pct", m12)}
        under5 = {"holders": cur["ff_holders"], "pct": cur["ff_pct"], "shares": cur["ff_shares"],
                  "holders_3m": grow("ff_holders", m3), "holders_12m": grow("ff_holders", m12),
                  "holders_chg_3m": diff("ff_holders", m3), "holders_chg_12m": diff("ff_holders", m12),
                  "pct_chg_3m": diff("ff_pct", m3), "pct_chg_12m": diff("ff_pct", m12),
                  "avg_shares": (cur["ff_shares"] / cur["ff_holders"]) if cur.get("ff_shares") and cur.get("ff_holders") else None}
        out[bank] = {
            "as_of": cur["date"], "total_shares": int(cur["total"]) if cur["total"] else None,
            "state_pct": cur["state_pct"], "over5": over5, "under5": under5,
            # хуучин кодтой нийцүүлэх (бүлэг бүрийн утга — нийлбэр БИШ)
            "dom_pct": cur["dom_pct"], "dom_holders": cur["dom_holders"], "dom_shares": cur["dom_shares"],
            "ff_pct": cur["ff_pct"], "ff_shares": cur["ff_shares"], "ff_holders": cur["ff_holders"],
            "qend_as_of": q["date"] if q else None,
            "qend_under5_holders": q["ff_holders"] if q else None, "qend_over5_holders": q["dom_holders"] if q else None,
            "history": [{"date": s["date"], "over5_holders": s["dom_holders"], "over5_pct": s["dom_pct"],
                         "under5_holders": s["ff_holders"], "under5_pct": s["ff_pct"]} for s in ss],
            "events": events,
        }
    return out


def load(url: str | None = DEFAULT_URL, cache: Path = DEFAULT_CACHE, file: str | None = None,
         quarter_end_iso: str | None = None) -> tuple[dict, str]:
    """Pipeline-ийн нэвтрэх цэг: (хураангуй, төлөвийн тайлбар). Алдаа гарвал ({}, шалтгаан)."""
    try:
        if file:
            path, status = Path(file), "локал файл"
        else:
            path, status = fetch(url, Path(cache))
        snaps = load_snapshots(path)
        summ = summarize(snaps, quarter_end_iso)
        latest = max((v["as_of"] for v in summ.values()), default="—")
        return summ, f"ҮЦТХТ төвлөрлийн судалгаа {latest} ({status})"
    except Exception as e:  # noqa: BLE001
        return {}, f"эзэмшлийн мэдээлэл ороогүй: {e}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default=DEFAULT_URL)
    ap.add_argument("--cache", default=str(DEFAULT_CACHE))
    ap.add_argument("--file", default=None)
    ap.add_argument("--quarter-end", default=None)
    a = ap.parse_args()
    summ, status = load(a.url, Path(a.cache), a.file, a.quarter_end)
    print(status)
    for b, s in summ.items():
        o, u = s["over5"], s["under5"]
        g3 = f"{u['holders_3m']:+.1%}" if u["holders_3m"] is not None else "—"
        g12 = f"{u['holders_12m']:+.1%}" if u["holders_12m"] is not None else "—"
        print(f"{b:11} {s['as_of']}  хувьцаа {s['total_shares']:>15,}")
        print(f"   5%-иас ДЭЭШ: {o['holders']:>3.0f} эзэмшигч, {o['pct']:.2%}, {o['shares']:>15,.0f} ш")
        print(f"   5%-иас ДООШ: {u['holders']:>7,.0f} эзэмшигч (3с {g3}, 12с {g12}), {u['pct']:.2%}, {u['shares']:>15,.0f} ш")
        for e in s["events"]:
            print(f"      {e['date']}: {e['text']}")
    print(json.dumps({"ok": bool(summ)}))


if __name__ == "__main__":
    main()
