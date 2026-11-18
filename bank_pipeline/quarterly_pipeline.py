#!/usr/bin/env python3
"""
quarterly_pipeline.py — Шинэ улирлын тайлан гарахад БҮГДИЙГ автоматаар хийнэ.

    1. ИЛРҮҮЛЭХ  — open.mse.mn дээр 5 банкны сүүлийн тайлангийн улирлыг workbook-тэй
                   харьцуулна. Голомт МХБ-д хоцорвол банкны IR сайтаас автоматаар авна.
    2. ТАТАХ     — тухайн тайлангийн БЛОК доторх БҮХ .xls/.xlsx хавсралтыг татаж,
                   reports/<YYYY>Q<n>/-д sha256-тай архивлана (Q4-д аудитын болон
                   аудитгүй хоёр файл байж болно — хамгийн бүрэн задлагдсаныг сонгоно).
    3. БИЧИХ     — bank_financials_bot.py-ийн баталгаажсан функцээр шинэ багана нэмнэ.
    4. ШАЛГАХ    — банк бүрт чанарын хаалт: Хөрөнгө = Өр + ӨХ, гол мөрүүд бичигдсэн,
                   мөрийн хамрах хүрээ, улирлын өөрчлөлт бодитой эсэх. Унавал тухайн
                   банкийг БУЦААНА (workbook-т хагас дутуу өгөгдөл үлдэхгүй).
    5. ДАХИН БОДОХ — LibreOffice recalc, томьёоны шинэ алдаа гарвал master-ийг солихгүй.
    6. ШИНЖИЛГЭЭ — bank_analysis.py: metrics, Word тайлан, график, fundamentals.json.
    7. МЭДЭГДЭХ  — Telegram-д дүгнэлт + Word тайлан.

Нэмэлт:
    • Нэг банк өнөөдөр, нөгөө нь 2 долоо хоногийн дараа гарсан ч болно — өдөр бүр
      ажиллуулахад шинэ гарсныг нь л нэмнэ; 5 банк бүгд ижил огноонд хүрэхэд
      гарчгийн огноог (Дашборд г.м.) автоматаар шинэчилнэ.
    • Q4-ийн аудитлагдсан тайлан (2-3 сарын дараа) ижил улирлын файлыг СОЛИВОЛ
      (sha өөрчлөгдвөл) тухайн баганыг дахин бичнэ (restatement).
    • --inbox DIR: open.mse.mn ажиллахгүй үед гараар татсан файлаа хавтаст хийгээд
      ажиллуулна (файлын нэр 564_..., SBM_..., эсвэл банкны нэрээр эхэлнэ).

Ажиллуулах:
    python3 quarterly_pipeline.py --workbook docs/Есөн_банк_шинжилгээ.xlsx
    python3 quarterly_pipeline.py --workbook W.xlsx --detect-only        # юу шинэ байгааг л хэвлэнэ
    python3 quarterly_pipeline.py --workbook W.xlsx --inbox inbox/ --period 2026Q4
    python3 quarterly_pipeline.py --workbook W.xlsx --analysis-only      # зөвхөн шинжилгээ
Орчны хувьсагч: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID (сонголт)
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402
import bank_analysis  # noqa: E402
import capitron_source  # noqa: E402
import transbank_source  # noqa: E402
import bogdbank_source  # noqa: E402

# МХБ-ийн бус банкууд — өөрсдийн сайтаас (bot.BANKS-д ороогүй)
EXTRA_SOURCES = {"Капитрон": capitron_source, "Транс банк": transbank_source, "Богд банк": bogdbank_source}
for _b in EXTRA_SOURCES:
    bot.HEADER_ROW.setdefault(_b, {"bal": 4, "inc": 4})
ALL_BANK_SHEETS = list(bot.BANK_TREND_OFFSET) + ["Капитрон", "Ариг банк", "Богд банк", "Транс банк"]

TICKER = {b: i["ticker"] for b, i in bot.BANKS.items()}
TICKER.update({b: m.TICKER for b, m in EXTRA_SOURCES.items()})
QDATE = {1: "03/31", 2: "06/30", 3: "09/30", 4: "12/31"}


def log(msg: str):
    print(msg, flush=True)


def period_of(date_label: str) -> tuple[int, int]:
    d = date_label.replace("*", "").strip()
    return int(d[:4]), int(d[5:7]) // 3


def pkey(p: tuple[int, int]) -> str:
    return f"{p[0]}Q{p[1]}"


# ═════════════════════════ 1. ИЛРҮҮЛЭХ / ТАТАХ ═════════════════════════════
XLS_RE = re.compile(r"\.xlsx?(?:$|\?)", re.I)


def discover_mse(company_id: int) -> dict | None:
    """Сүүлийн улирал + ТУХАЙН блокийн хавсралтууд. Хамгийн сүүлийн (он, улирал)-тай
    БҮХ блокийг цуглуулна (Q4-д аудитын/аудитгүй тайлан тусдаа блок байж болно)."""
    url = f"{bot.BASE_URL}/securities/{company_id}/tab/financials"
    resp = bot.safe_get(url)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    blocks = []
    for blk in soup.select(".finance-report-item"):
        h = blk.find("h5")
        text = h.get_text(" ", strip=True) if h else blk.get_text(" ", strip=True)[:120]
        m = re.search(r"(\d{4})\s*Он\s*(\d)\s*Улирал", text, re.I)
        if not m:
            continue
        links = [a["href"] for a in blk.find_all("a", href=True) if XLS_RE.search(a["href"])]
        blocks.append(((int(m.group(1)), int(m.group(2))), text, links))
    if not blocks:
        return None
    top = max(b[0] for b in blocks)
    links, heads = [], []
    for p, text, ls in blocks:
        if p == top:
            heads.append(text)
            links += [l if l.startswith("http") else bot.BASE_URL + l for l in ls if l not in links]
    if not links:
        # Блок дотор холбоос олдохгүй бол (хуудасны бүтэц өөрчлөгдсөн) — хуучин аргаар,
        # хуудасны эхний .xls холбоос. Огнооны шалгалт энэ эрсдэлийг барина.
        all_links = re.findall(r'href=["\']([^"\']*\.xlsx?)["\']', resp.text, re.I)
        links = [(l if l.startswith("http") else bot.BASE_URL + l) for l in all_links[:1]]
        log("  ⚠ Блок доторх холбоос олдсонгүй — хуудасны эхний холбоосыг авав (огноогоор шалгана)")
    return {"period": top, "heading": " | ".join(heads), "links": links}


def archive_bytes(content: bytes, ext: str, period, bank: str, root: Path) -> tuple[Path, str]:
    sha = hashlib.sha256(content).hexdigest()
    d = root / pkey(period)
    d.mkdir(parents=True, exist_ok=True)
    p = d / f"{TICKER.get(bank, bank)}_{sha[:10]}{ext}"
    if not p.exists():
        p.write_bytes(content)
    return p, sha


def parse_file(path: Path) -> dict[str, float]:
    raw: dict[str, float] = {}
    xls = pd.ExcelFile(path)
    for sh in xls.sheet_names:
        bot.extract_sheet_labels(pd.read_excel(xls, sheet_name=sh, header=None), raw)
    bot.aggregate_npl_categories(raw)
    return raw


DATE_PATTERNS = [
    r"(20\d\d)\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{1,2})",
    r"(20\d\d)\s*оны\s*(\d{1,2})\s*(?:дугаар|дүгээр|-?р)?\s*сарын\s*(\d{1,2})",
]


def file_dates(path: Path) -> set[tuple[int, int]]:
    """Файлын текстэнд байгаа улирлын эцсийн огноонууд -> {(он, улирал)}."""
    found = set()
    try:
        xls = pd.ExcelFile(path)
        for sh in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sh, header=None, nrows=15)
            txt = " ".join(str(v) for v in df.values.ravel() if isinstance(v, str))
            for v in df.values.ravel():
                if hasattr(v, "year") and hasattr(v, "month"):
                    txt += f" {v.year}/{v.month}/{getattr(v, 'day', 1)}"
            for pat in DATE_PATTERNS:
                for y, mth, d in re.findall(pat, txt, re.I):
                    mth, d = int(mth), int(d)
                    if (mth, d) in ((3, 31), (6, 30), (9, 30), (12, 31)):
                        found.add((int(y), mth // 3))
    except Exception:  # noqa: BLE001
        pass
    return found


CORE_HINTS = ["нийт хөрөнг", "өр төлбөр", "хадгаламж", "зээл", "хүүгийн орлого", "татварын дараах", "цэвэр ашиг"]


def score_parse(raw: dict[str, float]) -> tuple[int, int]:
    keys = " | ".join(bot.normalize_row_label(k) for k in raw)
    return sum(1 for h in CORE_HINTS if h in keys), len(raw)


def pick_best(cands: list[dict], period) -> dict | None:
    """Олон файлаас: огноо таарсан > гол шошго бүрэн > олон шошготой."""
    best = None
    for c in cands:
        dates = c["dates"]
        date_ok = 2 if period in dates else (1 if not dates else 0)
        c["date_ok"] = date_ok
        key = (date_ok,) + score_parse(c["raw"])[:1] + (1 if c.get("audited") else 0,) + score_parse(c["raw"])[1:]
        c["score"] = key
        if best is None or key > best["score"]:
            best = c
    return best


def fetch_candidates(bank: str, info: dict, period, links, reports_dir: Path) -> list[dict]:
    out = []
    for url in links:
        try:
            r = bot.safe_get(url, timeout=60)
            r.raise_for_status()
        except Exception as e:  # noqa: BLE001
            log(f"  ⚠ татаж чадсангүй {url}: {e}")
            continue
        ext = ".xlsx" if url.lower().split("?")[0].endswith(".xlsx") else ".xls"
        path, sha = archive_bytes(r.content, ext, period, bank, reports_dir)
        try:
            raw = parse_file(path)
        except Exception as e:  # noqa: BLE001
            log(f"  ⚠ задалж чадсангүй {path.name}: {e}")
            continue
        out.append({"url": url, "path": path, "sha": sha, "raw": raw, "dates": file_dates(path)})
        log(f"  · {path.name}: {len(raw)} шошго, файлын огноо {sorted(out[-1]['dates'])}")
    return out


def inbox_candidates(bank: str, info: dict, period, inbox: Path, reports_dir: Path) -> list[dict]:
    out = []
    keys = [k for k in (str(info["id"]) if info.get("id") else None, info["ticker"].lower(), bank.lower()) if k]
    for f in sorted(inbox.iterdir()):
        if f.suffix.lower() not in (".xls", ".xlsx", ".pdf", ".html") or not any(f.name.lower().startswith(k) for k in keys):
            continue
        path, sha = archive_bytes(f.read_bytes(), f.suffix.lower(), period, bank, reports_dir)
        if bank in EXTRA_SOURCES:
            raw = EXTRA_SOURCES[bank].adapt_to_sheet(EXTRA_SOURCES[bank].parse_file(path))
        else:
            raw = parse_file(path)
        dates = file_dates(path) if path.suffix.lower() in (".xls", ".xlsx") else set()
        out.append({"url": f"inbox:{f.name}", "path": path, "sha": sha, "raw": raw, "dates": dates})
        log(f"  · inbox {f.name}: {len(raw)} шошго, файлын огноо {sorted(out[-1]['dates'])}")
    return out


def golomt_site_candidate(period=None) -> dict | None:
    try:
        res = bot.fetch_golomt_site(f"{period[0]}/{QDATE[period[1]]}" if period else None)
    except Exception as e:  # noqa: BLE001
        log(f"  ⚠ Голомтын сайт: {e}")
        return None
    if not res:
        return None
    d, raw = res
    p = period_of(d)
    sha = hashlib.sha256(json.dumps(raw, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    return {"url": bot.GOLOMT_SITE_URL, "path": None, "sha": sha, "raw": raw, "dates": {p},
            "period": p, "date_ok": 2, "score": (2, 0, len(raw))}


# ═════════════════════════ 2. ЧАНАРЫН ХААЛТ ════════════════════════════════
def find_row(ws, hdr, names) -> int | None:
    names = {n.lower() for n in names}
    for r in range(hdr + 1, ws.max_row + 1):
        if re.sub(r"\s+", " ", str(ws.cell(r, 1).value or "")).strip().lower() in names:
            return r
    return None


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def quality_gates(ws_bal, hb, ws_inc, hi, col_b, col_i, prev_b, prev_i) -> list[str]:
    errs = []
    ra = find_row(ws_bal, hb, ["НИЙТ ХӨРӨНГИЙН ДҮН"])
    rl = find_row(ws_bal, hb, ["НИЙТ ӨР ТӨЛБӨРИЙН ДҮН", "НИЙТ ӨР ТӨЛБӨР"])
    re_ = find_row(ws_bal, hb, ["ӨӨРИЙН ХӨРӨНГИЙН ДҮН", "НИЙТ ӨӨРИЙН ХӨРӨНГӨ"])
    rp = find_row(ws_inc, hi, ["Татварын дараах ашиг"])
    a, l, e = (num(ws_bal.cell(r, col_b).value) if r else None for r in (ra, rl, re_))
    p = num(ws_inc.cell(rp, col_i).value) if rp else None
    for name, v in (("Нийт хөрөнгө", a), ("Нийт өр төлбөр", l), ("Өөрийн хөрөнгө", e), ("Цэвэр ашиг", p)):
        if v is None:
            errs.append(f"'{name}' бичигдсэнгүй")
    if None not in (a, l, e) and abs(a - l - e) > max(1.0, 0.005 * abs(a)):
        errs.append(f"Хөрөнгө≠Өр+ӨХ (зөрүү {a - l - e:,.0f})")
    pa = num(ws_bal.cell(ra, prev_b).value) if ra else None
    if a and pa and not (0.7 <= a / pa <= 1.4):
        errs.append(f"Нийт хөрөнгө өмнөх улирлаас {a / pa - 1:+.0%} — масштаб/мөрийн алдаа байж болзошгүй")
    for ws, hdr, c, pc, nm in ((ws_bal, hb, col_b, prev_b, "баланс"), (ws_inc, hi, col_i, prev_i, "орлого")):
        had = sum(1 for r in range(hdr + 1, ws.max_row + 1) if num(ws.cell(r, pc).value) is not None)
        got = sum(1 for r in range(hdr + 1, ws.max_row + 1) if num(ws.cell(r, c).value) is not None)
        if had and got < 0.6 * had:
            errs.append(f"{nm}: {got}/{had} мөр л бичигдсэн")
    return errs


def snapshot_col(ws, col):
    return {r: ws.cell(r, col).value for r in range(1, ws.max_row + 1)}


def restore_col(ws, col, snap):
    for r in range(1, ws.max_row + 1):
        ws.cell(r, col).value = snap.get(r)


def find_date_col(ws, hdr, date_label) -> int | None:
    t = date_label.replace("*", "")
    for c in range(2, ws.max_column + 1):
        if str(ws.cell(hdr, c).value or "").replace("*", "").strip() == t:
            return c
    return None


# ═════════════════════════ 2b. ЕРӨНХИЙ (9 банкны аль ч) ДООД УРСГАЛ ═══════
def add_trend_columns_generic(wb, bank, date_label, old_b, new_b, old_i, new_i) -> int:
    """bot.add_trend_columns нь зөвхөн МХБ-ийн 5 банкны мөрийн offset-оор ажилладаг.
    Энэ хувилбар банкны НЭРЭЭР мөрийг олж (A багана), тухайн мөрийн хамгийн сүүлийн
    томьёог хуулж баганын үсгийг солино. Шинэ огнооны багана аль хэдийн байвал
    (5 банк түрүүлж оруулсан) тэр баганад бичнэ; блок бүрийн "Банк / Улирал"
    толгойг бүгдийг нь бөглөнө."""
    label = bot.trend_date_label(date_label)
    ob, nb, oi, ni = map(bot.col_letter, (old_b, new_b, old_i, new_i))
    n = 0
    for sheet in bot.TREND_BLOCKS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        hdr_rows = [r for r in range(1, ws.max_row + 1)
                    if str(ws.cell(r, 1).value or "").strip().lower().startswith("банк / улирал")]
        if not hdr_rows:
            continue
        target = None
        for c in range(2, ws.max_column + 1):
            if any(str(ws.cell(h, c).value or "").strip() == label for h in hdr_rows):
                target = c
                break
        if target is None:
            target = ws.max_column + 1
        for r in range(1, ws.max_row + 1):
            if str(ws.cell(r, 1).value or "").strip() != bank:
                continue
            src_c = next((c for c in range(target - 1, 1, -1)
                          if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.startswith("=")), None)
            if src_c is None:
                continue
            f = ws.cell(r, src_c).value
            f2 = re.sub(rf"('{re.escape(bank)} баланс'!){ob}(\d+)", rf"\g<1>{nb}\2", f)
            f2 = re.sub(rf"('{re.escape(bank)} орлого'!){oi}(\d+)", rf"\g<1>{ni}\2", f2)
            if f2 == f:
                continue
            cell = ws.cell(r, target, value=f2)
            bot.copy_cell_style(ws.cell(r, src_c), cell)
            n += 1
        for h in hdr_rows:
            if ws.cell(h, target).value in (None, ""):
                cell = ws.cell(h, target, value=label)
                bot.copy_cell_style(ws.cell(h, target - 1), cell)
    return n


def update_annualization_generic(wb, bank: str, quarter: int) -> int:
    """bot.update_annualization-тай ижил, гэхдээ 9 банкны аль ч нэрээр эзэмшигчийг олно."""
    factor = bot.ANNUAL_FACTOR[quarter]
    n = 0
    for sheet in bot.ANNUALIZE_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        for row in wb[sheet].iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=") and bot._ANN_RE.search(v)):
                    continue
                owners = [b for b in ALL_BANK_SHEETS if f"'{b} " in v]
                if owners[:1] != [bank]:
                    continue
                nv = bot._ANN_RE.sub(f"*{factor}", v)
                if nv != v:
                    cell.value = nv
                    n += 1
    return n


# ═════════════════════════ 3. БИЧИХ ════════════════════════════════════════
def ingest(wb, bank: str, period, raw: dict, restate: bool = False) -> tuple[bool, str]:
    y, q = period
    date_label = bot.quarter_end_date(y, q)
    ws_bal, ws_inc = wb[f"{bank} баланс"], wb[f"{bank} орлого"]
    hb, hi = bot.HEADER_ROW[bank]["bal"], bot.HEADER_ROW[bank]["inc"]
    _, last_b = bot.get_excel_last_date(ws_bal, hb)
    _, last_i = bot.get_excel_last_date(ws_inc, hi)
    cum = bank == "ХасБанк"

    if restate:
        col_b, col_i = find_date_col(ws_bal, hb, date_label), find_date_col(ws_inc, hi, date_label)
        if col_b != last_b or col_i != last_i:
            return False, "дахин бичих багана сүүлийн багана биш (дараагийн улирал орсон) — гараар шийднэ"
        snap_b, snap_i = snapshot_col(ws_bal, col_b), snapshot_col(ws_inc, col_i)
        for ws, c, h in ((ws_bal, col_b, hb), (ws_inc, col_i, hi)):
            for r in range(h + 1, ws.max_row + 1):
                if num(ws.cell(r, c).value) is not None:
                    ws.cell(r, c).value = None
        wb_, eb = bot.write_new_quarter_full(ws_bal, hb, col_b, date_label, raw)
        wi_, ei = bot.write_new_quarter_full(ws_inc, hi, col_i, date_label, raw, cumulative_flow=cum)
        if cum:
            bot.fix_khasbank_cumulative_to_standalone(ws_inc, hi, col_i, date_label)
        errs = quality_gates(ws_bal, hb, ws_inc, hi, col_b, col_i, col_b - 1, col_i - 1)
        if errs:
            restore_col(ws_bal, col_b, snap_b)
            restore_col(ws_inc, col_i, snap_i)
            return False, "; ".join(errs)
        return True, f"{date_label} дахин бичигдлээ (баланс {len(wb_)}, орлого {len(wi_)} мөр)"

    col_b, col_i = last_b + 1, last_i + 1
    wb_, eb = bot.write_new_quarter_full(ws_bal, hb, col_b, date_label, raw)
    wi_, ei = bot.write_new_quarter_full(ws_inc, hi, col_i, date_label, raw, cumulative_flow=cum)
    if cum:
        bot.fix_khasbank_cumulative_to_standalone(ws_inc, hi, col_i, date_label)
    errs = quality_gates(ws_bal, hb, ws_inc, hi, col_b, col_i, last_b, last_i)
    if errs:
        for ws, c in ((ws_bal, col_b), (ws_inc, col_i)):
            for r in range(1, ws.max_row + 1):
                ws.cell(r, c).value = None
        return False, "; ".join(errs)

    bot.repoint_snapshot_formulas(wb, bank, last_b, col_b, last_i, col_i)
    if cum:
        bot.rebuild_khasbank_ytd(wb, date_label)
    if bank in bot.BANK_TREND_OFFSET:
        bot.update_annualization(wb, bank, q)
        bot.add_trend_columns(wb, bank, date_label, last_b, col_b, last_i, col_i)
    else:
        update_annualization_generic(wb, bank, q)
        add_trend_columns_generic(wb, bank, date_label, last_b, col_b, last_i, col_i)
    empty = [x for x in eb + ei if x not in bot.SECTION_HEADER_LABELS]
    msg = f"баланс {len(wb_)}, орлого {len(wi_)} мөр"
    if empty:
        msg += f"; хоосон үлдсэн: {', '.join(empty[:6])}{'…' if len(empty) > 6 else ''}"
    return True, msg


def sync_period_labels(wb) -> int:
    """5 банк бүгд ижил сүүлийн огноонд хүрсэн ба гарчигт хуучин огноо байвал
    шинэчилнэ. Банкууд ӨӨР өдрүүдэд орж ирсэн ч ажиллана. МХБ-ийн бус 4 банкны
    огноог ("Капитрон, Ариг, …— 2026/06/30") ХЭЗЭЭ Ч өөрчлөхгүй — тэр хэсгийг
    түр тасдаж аваад, шинэчлэлийн дараа буцааж залгана."""
    dates = set()
    for bank in bot.BANKS:
        d, _ = bot.get_excel_last_date(wb[f"{bank} баланс"], bot.HEADER_ROW[bank]["bal"])
        dates.add(d.replace("*", ""))
    if len(dates) != 1:
        return 0
    new = dates.pop()
    marker = bot.NON_MSE_BANKS.split(",")[0]          # "Капитрон"
    tails = {}
    seen = []
    for sheet, cells in bot.PERIOD_LABEL_CELLS.items():
        if sheet not in wb.sheetnames:
            continue
        for c in cells:
            v = wb[sheet][c].value
            if not isinstance(v, str) or v.startswith("="):
                continue
            i = v.find(marker)
            if i > 0:
                tails[(sheet, c)] = v[i:]
                wb[sheet][c].value = v[:i]
                v = v[:i]
            ds = re.findall(r"20\d\d/\d\d/\d\d", v)
            if ds and new not in ds:          # шинэ огноо аль хэдийн байгаа нүдийг хөндөхгүй
                seen += ds
    olds = [x for x in seen if x < new]
    n = 0
    if olds:
        # Зөвхөн ХАМГИЙН СҮҮЛИЙН хуучин огноо (= өмнөх улирал); "2023/12/31 — …" хөндөгдөхгүй
        n = bot.update_period_labels(wb, max(olds), new, all_five_updated=True)
    for (sheet, c), tail in tails.items():
        wb[sheet][c].value = (wb[sheet][c].value or "") + tail
    return n


def refresh_non_mse_tail(wb) -> int:
    """Гарчгийн "Капитрон, Ариг, Богд, Транс банк — 2026/06/30" хэсгийг 4 банкны
    БОДИТ сүүлийн огноогоор дахин бичнэ (огноогоор бүлэглэнэ). Жиш.:
    "Капитрон, Ариг банк — 2026/06/30; Богд, Транс банк — 2026/09/30"."""
    order = ["Капитрон", "Ариг банк", "Богд банк", "Транс банк"]
    dates = {}
    for b in order:
        if f"{b} баланс" not in wb.sheetnames:
            continue
        ws = wb[f"{b} баланс"]
        hdr = bot.HEADER_ROW.get(b, {}).get("bal", 4)
        d, _ = bot.get_excel_last_date(ws, hdr)
        dates[b] = d.replace("*", "")
    groups = {}
    for b in order:
        if b in dates:
            groups.setdefault(dates[b], []).append(b)
    def names(bs):
        short = [x.replace(" банк", "") for x in bs]
        return ", ".join(short[:-1] + [bs[-1]]) if len(bs) > 1 else bs[0]
    text = "; ".join(f"{names(bs)} — {d}" for d, bs in sorted(groups.items()))
    n = 0
    pat = re.compile(r"Капитрон[^;.]*?—\s*20\d\d/\d\d/\d\d(?:;\s*[^;.]*?—\s*20\d\d/\d\d/\d\d)*")
    for sheet, cells in bot.PERIOD_LABEL_CELLS.items():
        if sheet not in wb.sheetnames:
            continue
        for c in cells:
            v = wb[sheet][c].value
            if isinstance(v, str) and not v.startswith("=") and pat.search(v):
                nv = pat.sub(text, v, count=1)
                if nv != v:
                    wb[sheet][c].value = nv
                    n += 1
    return n


def formula_errors(path: str) -> int:
    wv = load_workbook(path, data_only=True)
    n = 0
    for ws in wv.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip().upper() in (
                        "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"):
                    n += 1
    return n


# ═════════════════════════ 4. TELEGRAM ════════════════════════════════════
def tg_send_document(document: Path | None) -> bool:
    """Telegram-д ЗӨВХӨН шинжилгээний Word файлыг (тайлбаргүй) илгээнэ.
    Тайлан орсон төлөв, алдаа зэрэг нь Word тайлангийн "Өгөгдлийн төлөв" бүлэгт байна."""
    tok, chat = os.environ.get("TELEGRAM_BOT_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
    if not document or not Path(document).exists():
        log("ℹ Telegram: илгээх Word файл алга")
        return False
    if not tok or not chat:
        log(f"ℹ Telegram тохируулаагүй — илгээх байсан файл: {Path(document).name}")
        return False
    import requests
    for attempt in (1, 2):
        try:
            with open(document, "rb") as f:
                r = requests.post(f"https://api.telegram.org/bot{tok}/sendDocument", data={"chat_id": chat},
                                  files={"document": (Path(document).name, f)}, timeout=120)
            if r.status_code < 400:
                log(f"✓ Telegram: {Path(document).name} илгээгдлээ")
                return True
            log(f"⚠ Telegram HTTP {r.status_code}: {getattr(r, 'text', '')[:200]}")
        except Exception as e:  # noqa: BLE001
            log(f"⚠ Telegram (оролдлого {attempt}): {e}")
    return False


# ═════════════════════════ 5. ҮНДСЭН УРСГАЛ ════════════════════════════════
def own_kw(a) -> dict:
    """Эзэмшлийн (ҮЦТХТ) мэдээллийг шинжилгээнд дамжуулах тохиргоо."""
    adv = {"adv_params": {"coe": a.coe, "growth": a.growth, "lgd": a.lgd}}
    if a.no_ownership:
        return {"ownership": {}, "ownership_note": "эзэмшлийн мэдээлэл унтраалттай (--no-ownership)", **adv}
    return {"ownership_url": a.ownership_url, "ownership_cache": a.ownership_cache,
            "ownership_file": a.ownership_file, **adv}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--reports-dir", default="reports")
    ap.add_argument("--state", default="state/report_state.json")
    ap.add_argument("--analysis-dir", default="analysis")
    ap.add_argument("--archive-dir", default="archive")
    ap.add_argument("--prices", default=None)
    ap.add_argument("--fundamentals", default=None)
    ap.add_argument("--inbox", default=None, help="гараар татсан тайлангийн хавтас")
    ap.add_argument("--period", default=None, help="--inbox-д: 2026Q4 гэх мэт")
    ap.add_argument("--detect-only", action="store_true")
    ap.add_argument("--analysis-only", action="store_true")
    ap.add_argument("--no-restate", action="store_true")
    ap.add_argument("--ownership-url", default=None,
                    help="ҮЦТХТ төвлөрлийн xlsx (анхдагч: consultingapartment/Huvitsaanii-Tuvlurul)")
    ap.add_argument("--ownership-cache", default="data/ownership/uctht_concentration.xlsx")
    ap.add_argument("--ownership-file", default=None, help="локал файл (татахгүй)")
    ap.add_argument("--no-ownership", action="store_true")
    ap.add_argument("--coe", type=float, default=None, help="үндэслэлтэй P/B: хувьцааны өгөөжийн шаардлага")
    ap.add_argument("--growth", type=float, default=None, help="үндэслэлтэй P/B: урт хугацааны өсөлт")
    ap.add_argument("--lgd", type=float, default=None, help="стресс тест: шинэ NPL-ийн алдагдлын хувь")
    ap.add_argument("--force-analysis", action="store_true")
    ap.add_argument("--github-output", default=os.environ.get("GITHUB_OUTPUT"))
    a = ap.parse_args(argv)

    master = Path(a.workbook)
    reports_dir = Path(a.reports_dir)
    state_p = Path(a.state)
    state = json.loads(state_p.read_text(encoding="utf-8")) if state_p.exists() else {}

    if a.analysis_only:
        res = bank_analysis.run(str(master), a.analysis_dir, a.prices, a.fundamentals, **own_kw(a))
        tg_send_document(res["docx"])
        return 0

    wb = load_workbook(master)
    plan = []          # (bank, period, candidate, restate)
    notes = []
    banks_iter = list(bot.BANKS.items()) + [(b, {"id": None, "ticker": m.TICKER}) for b, m in EXTRA_SOURCES.items()]
    for bank, info in banks_iter:
        log(f"\n── {bank} ({info['ticker']})")
        ws_bal = wb[f"{bank} баланс"]
        excel_date, _ = bot.get_excel_last_date(ws_bal, bot.HEADER_ROW[bank]["bal"])
        excel_p = period_of(excel_date)
        cands, src_p = [], None

        if a.inbox:
            if not a.period:
                sys.exit("--inbox ашиглахад --period (жиш. 2026Q4) заавал")
            y, q = map(int, a.period.upper().split("Q"))
            src_p = (y, q)
            cands = inbox_candidates(bank, info, src_p, Path(a.inbox), reports_dir)
        elif bank in EXTRA_SOURCES:
            try:
                src_p, cands, heading = EXTRA_SOURCES[bank].candidates_for(excel_p, reports_dir)
                log(f"  Сайт: {heading} → {pkey(src_p) if src_p else '—'}; workbook: {excel_date}")
            except Exception as e:  # noqa: BLE001
                src_p, cands = None, []
                log(f"  ⚠ {bank} сайт: {e}")
        else:
            try:
                disc = discover_mse(info["id"])
            except Exception as e:  # noqa: BLE001
                disc = None
                log(f"  ⚠ МХБ: {e}")
            if disc:
                src_p = disc["period"]
                log(f"  МХБ: {disc['heading']} → {pkey(src_p)}; workbook: {excel_date}; {len(disc['links'])} файл")
                if src_p >= excel_p:   # ижил улирал ч татна — эх файл солигдсон эсэхийг (sha) шалгана
                    cands = fetch_candidates(bank, info, src_p, disc["links"], reports_dir)
            if bank == "Голомт" and (src_p is None or src_p <= excel_p or not cands):
                site = golomt_site_candidate(None if (src_p is None or src_p <= excel_p) else src_p)
                if site and site["period"] > excel_p:
                    log(f"  Голомтын IR сайт: {pkey(site['period'])} ({len(site['raw'])} шошго)")
                    src_p, cands = site["period"], [site]

        if src_p is None:
            notes.append(f"⚠️ {bank}: тайлангийн мэдээлэл олдсонгүй")
            continue
        if src_p < excel_p or (src_p == excel_p and a.no_restate):
            log(f"  ✓ шинэ зүйлгүй ({excel_date})")
            continue

        best = pick_best([c for c in cands if c.get("raw")], src_p) if cands else None
        if best is None:
            if src_p > excel_p:
                notes.append(f"❌ {bank}: {pkey(src_p)} тайлан гарсан ч файл татагдсангүй/задлагдсангүй")
            continue
        if best.get("date_ok") == 0:
            notes.append(f"❌ {bank}: файлын огноо {sorted(best['dates'])} нь {pkey(src_p)}-тэй таарахгүй — бичсэнгүй, "
                         f"{best['path']} -г шалгана уу")
            continue

        if src_p == excel_p:   # ижил улирал — файл солигдсон эсэх (аудит/засвар)
            prev_sha = state.get(bank, {}).get(pkey(src_p), {}).get("sha")
            if prev_sha is None:
                state.setdefault(bank, {})[pkey(src_p)] = {"sha": best["sha"], "url": best["url"],
                                                          "path": str(best["path"]) if best["path"] else None,
                                                          "baseline": True}
                log("  ✓ шинэ зүйлгүй (эх файлын hash суурь болгон хадгалав)")
                continue
            if prev_sha == best["sha"]:
                log(f"  ✓ шинэ зүйлгүй ({excel_date})")
                continue
            log("  ♻ ижил улирлын эх файл СОЛИГДСОН — дахин бичнэ")
            plan.append((bank, src_p, best, True))
        else:
            plan.append((bank, src_p, best, False))

    gh = {}
    if a.detect_only:
        gh["new"] = "true" if plan else "false"
        if a.github_output:
            with open(a.github_output, "a") as f:
                for k, v in gh.items():
                    f.write(f"{k}={v}\n")
        for bank, p, _, rs in plan:
            log(f"🆕 {bank}: {pkey(p)}{' (эх файл солигдсон)' if rs else ''}")
        log(f"\nИлрүүлэлт: {len(plan)} банкинд шинэ/солигдсон тайлан")
        return 0

    results = []
    for bank, p, cand, restate in plan:
        ok, msg = ingest(wb, bank, p, cand["raw"], restate=restate)
        tag = "♻" if restate else "🆕"
        log(f"{'✓' if ok else '❌'} {bank} {pkey(p)}: {msg}")
        results.append((bank, p, ok, msg, tag, cand))
        if ok:
            state.setdefault(bank, {})[pkey(p)] = {"sha": cand["sha"], "url": cand["url"],
                                                  "path": str(cand["path"]) if cand["path"] else None,
                                                  "ingested": datetime.now().isoformat(timespec="seconds")}
            if cand.get("path"):
                (cand["path"].with_suffix(".labels.json")).write_text(
                    json.dumps(cand["raw"], ensure_ascii=False, indent=1), encoding="utf-8")

    any_ok = any(r[2] for r in results)
    lines = [f"🏦 <b>Банкны улирлын тайлан</b> — {datetime.now():%Y-%m-%d %H:%M}"]
    for bank, p, ok, msg, tag, _ in results:
        lines.append(f"{tag if ok else '❌'} <b>{bank}</b> {pkey(p)}: {msg}")
    lines += notes

    if any_ok:
        n_lbl = sync_period_labels(wb) + refresh_non_mse_tail(wb)
        if n_lbl:
            log(f"✓ Гарчгийн огноо: {n_lbl} нүд")
        tmp = Path(tempfile.mkdtemp()) / master.name
        wb.save(tmp)
        bot.recalc_workbook(str(tmp), timeout=240)
        err_before, err_after = formula_errors(str(master)), formula_errors(str(tmp))
        if err_after > err_before:
            cand_path = master.with_name(master.stem + "_CANDIDATE.xlsx")
            shutil.copy(tmp, cand_path)
            lines.append(f"⛔ Томьёоны алдаа {err_before}→{err_after} болж өссөн тул master солигдсонгүй. "
                         f"Шалгах файл: {cand_path.name}")
            log(lines[-1])
            state_p.parent.mkdir(parents=True, exist_ok=True)
            state_p.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
            return 2
        arch = Path(a.archive_dir)
        arch.mkdir(parents=True, exist_ok=True)
        shutil.copy(master, arch / f"{master.stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy(tmp, master)
        log(f"✓ {master} шинэчлэгдлээ (хуучин хувилбар {arch}/-д)")

    state_p.parent.mkdir(parents=True, exist_ok=True)
    state_p.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")

    if any_ok or a.force_analysis:
        res = bank_analysis.run(str(master), a.analysis_dir, a.prices, a.fundamentals, **own_kw(a),
                                status_lines=lines)
        tg_send_document(res["docx"])
        gh["analysis_dir"] = str(res["out_dir"])
    elif results or notes:
        log("\n".join(lines))   # шинжилгээ гараагүй тул Telegram-д юу ч явуулахгүй (лог-д үлдэнэ)
    else:
        log("\nℹ Шинэ тайлан алга — юу ч өөрчлөгдсөнгүй")
    gh["updated"] = "true" if any_ok else "false"
    if a.github_output:
        with open(a.github_output, "a") as f:
            for k, v in gh.items():
                f.write(f"{k}={v}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
