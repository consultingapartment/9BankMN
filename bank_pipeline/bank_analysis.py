#!/usr/bin/env python3
"""
bank_analysis.py — Есөн банкны workbook-оос улирлын шинжилгээг АВТОМАТААР гаргана.

Workbook-ийн томьёонд (Дашборд г.м.) найдахгүй — банк тус бүрийн баланс/орлого
хуудасны ТҮҮХИЙ утгаас шууд тооцоолно. Ингэснээр recalc хийгдээгүй, эсвэл
snapshot томьёо буруу багана заасан үед ч шинжилгээ зөв гарна.

Гаргах зүйлс (analysis/<YYYY>Q<n>/ хавтаст):
    metrics.json, metrics.csv     — 9 банкны бүх үзүүлэлт
    report.docx                   — Q2 тайлантай ижил бүтэцтэй шинжилгээний тайлан
    chart_*.png                   — графикууд
    telegram.txt                  — Telegram-д явуулах товч дүгнэлт
    + data/fundamentals.json      — dashboard-ийн EPS/BVPS/P/E/P/B (TTM)

Ажиллуулах:
    python3 bank_analysis.py <workbook.xlsx> [--out analysis] [--prices data/price_data.json]
                             [--fundamentals data/fundamentals.json]
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import statistics
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

BANKS = ["Төрийн банк", "ХасБанк", "ХХБ", "Голомт", "Хаан банк",
         "Капитрон", "Ариг банк", "Богд банк", "Транс банк"]
MSE_BANKS = ["Төрийн банк", "ХасБанк", "ХХБ", "Голомт", "Хаан банк"]
STANDALONE_INCOME = {"ХасБанк"}          # орлого улирлын ДАН дүнгээр хадгалагддаг
# Гэхдээ "ХасБанк орлого" хуудасны тэмдэглэлд дурдсанаар 2023/12–2024/12 багануудыг
# нэмэлт файлаас ХУРИМТЛАЛ (YTD)-аар нөхсөн — тэр хүртэлх үеийг хуримтлал гэж үзнэ.
STANDALONE_FROM = {"ХасБанк": (2025, 1)}
ANNUAL = {1: 4.0, 2: 2.0, 3: 4.0 / 3.0, 4: 1.0}
PERIOD_MN = {1: "1-р улирал (3 сар)", 2: "эхний хагас жил (H1)", 3: "9 сар", 4: "бүтэн жил"}
PERIOD_SHORT = {1: "3 сар", 2: "H1", 3: "9 сар", 4: "FY"}
QUARTER_ROMAN = {1: "I", 2: "II", 3: "III", 4: "IV"}

# Шошго (норчигжсон) -> үзүүлэлт. Эхний олдсоныг авна.
BAL_KEYS = {
    "assets": ["нийт хөрөнгийн дүн"],
    "liabilities": ["нийт өр төлбөрийн дүн", "нийт өр төлбөр"],
    "equity": ["өөрийн хөрөнгийн дүн", "нийт өөрийн хөрөнгө"],
    "loans_net": ["зээл (цэвэр дүнгээр)"],
    "loans_normal": ["хэвийн зээл"],
    "loans_watch": ["анхаарал хандуулах зээл", "хугацаа хэтэрсэн зээл"],
    "npl": ["чанаргүй зээл", "чанаргүй зээл (npl)"],
    "npl_sub": ["хэвийн бус зээл"],
    "npl_doubt": ["эргэлзээтэй зээл"],
    "npl_loss": ["муу зээл"],
    "provision": ["зээлийн эрсдэлийн сан"],
    "current_acc": ["харилцах"],
    "deposits_only": ["хадгаламж"],
    "deposits_combined": ["харилцагчдын харилцах, хадгаламж"],
    "cash": ["мөнгө ба түүнтэй адилтгах хөрөнгө"],
}
INC_KEYS = {
    "int_income": ["хүүгийн орлого"],
    "int_expense": ["хүүгийн зардал"],
    "nii": ["хүүгийн цэвэр орлого"],
    "provision_exp": ["эрсдэлийн сангийн зардал"],
    "nonint_income": ["хүүгийн бус орлого", "шимтгэл хураамж, бусад цэвэр орлого"],
    "nonint_expense": ["хүүгийн бус зардал", "үйл ажиллагааны зардал"],
    "pbt": ["татварын өмнөх ашиг"],
    "net_profit": ["татварын дараах ашиг"],
}


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def qkey(d: str) -> tuple[int, int]:
    """'2026/09/30' -> (2026, 3)"""
    d = d.replace("*", "")
    return int(d[:4]), int(d[5:7]) // 3


def qlabel(y: int, q: int) -> str:
    return f"{y}/{['03/31', '06/30', '09/30', '12/31'][q - 1]}"


def prev_q(y: int, q: int) -> tuple[int, int]:
    return (y, q - 1) if q > 1 else (y - 1, 4)


class BankData:
    """Нэг банкны баланс/орлогын (огноо -> үзүүлэлт) хүснэгт."""

    def __init__(self, wb, bank: str):
        self.bank = bank
        self.bal = self._read(wb[f"{bank} баланс"], BAL_KEYS)
        self.inc = self._read(wb[f"{bank} орлого"], INC_KEYS)

    @staticmethod
    def _header_row(ws) -> int:
        for r in range(1, 9):
            if norm(ws.cell(r, 1).value) == "үзүүлэлт":
                if any(re.match(r"\d{4}/\d{2}/\d{2}", str(ws.cell(r, c).value or "")) for c in range(2, ws.max_column + 1)):
                    return r
        raise ValueError(f"{ws.title}: огнооны толгой мөр олдсонгүй")

    def _read(self, ws, keys) -> dict[tuple[int, int], dict[str, float]]:
        hdr = self._header_row(ws)
        rows = {}
        for r in range(hdr + 1, ws.max_row + 1):
            n = norm(ws.cell(r, 1).value)
            if n and n not in rows:
                rows[n] = r
        out: dict[tuple[int, int], dict[str, float]] = {}
        for c in range(2, ws.max_column + 1):
            h = str(ws.cell(hdr, c).value or "").strip()
            if not re.match(r"\d{4}/\d{2}/\d{2}", h):
                continue
            vals = {}
            for key, labels in keys.items():
                for lab in labels:
                    r = rows.get(lab)
                    if r is not None:
                        v = ws.cell(r, c).value
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            vals[key] = float(v)
                            break
            if vals:
                out[qkey(h)] = vals
        return out

    # ── түвшин ────────────────────────────────────────────────────────────
    def latest(self) -> tuple[int, int] | None:
        common = [k for k in self.bal if k in self.inc and "assets" in self.bal[k]]
        return max(common) if common else None

    def b(self, k, key):
        return self.bal.get(k, {}).get(key)

    def _standalone(self, k) -> bool:
        return self.bank in STANDALONE_INCOME and k >= STANDALONE_FROM.get(self.bank, (0, 0))

    def ytd(self, k, key):
        """Оны эхнээс хуримтлал (YTD)."""
        if self._standalone(k):
            y, q = k
            parts = [self.inc.get((y, i), {}).get(key) for i in range(1, q + 1)]
            return sum(parts) if all(p is not None for p in parts) else None
        return self.inc.get(k, {}).get(key)

    def quarter_alone(self, k, key):
        """Улирлын ДАН дүн."""
        if self._standalone(k):
            return self.inc.get(k, {}).get(key)
        y, q = k
        cur = self.ytd(k, key)
        if cur is None:
            return None
        if q == 1:
            return cur
        prev = self.ytd((y, q - 1), key)
        return None if prev is None else cur - prev

    def ttm(self, k, key):
        y, q = k
        cur = self.ytd(k, key)
        if q == 4:
            return cur
        fy = self.ytd((y - 1, 4), key)
        prev_same = self.ytd((y - 1, q), key)
        if None in (cur, fy, prev_same):
            return None
        return cur + fy - prev_same

    def gross_loans_npl(self, k):
        """Workbook-ийн 'Активын чанар' хуудастай ИЖИЛ тодорхойлолт:
        нийт (бохир) зээл = цэвэр зээл + |эрсдэлийн сан|; NPL = чанаргүй
        (эсвэл хэвийн бус + эргэлзээтэй + муу)."""
        npl = self.b(k, "npl")
        if npl is None:
            parts = [self.b(k, x) for x in ("npl_sub", "npl_doubt", "npl_loss")]
            if all(p is not None for p in parts):
                npl = sum(parts)
        net, prov = self.b(k, "loans_net"), self.b(k, "provision")
        gross = net + abs(prov) if (net is not None and prov is not None) else None
        return gross, npl

    def deposits(self, k):
        comb = self.b(k, "deposits_combined")
        if comb is not None:
            return comb
        a, d = self.b(k, "current_acc"), self.b(k, "deposits_only")
        if a is None and d is None:
            return None
        # Богд банк "Харилцах, хадгаламж"-аа нэг мөрөнд (Харилцах) нэгтгэж тайлагнадаг
        return (a or 0.0) + (d or 0.0)


def pct(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


def growth(a, b):
    if a is None or b in (None, 0):
        return None
    return (a - b) / abs(b)   # сөрөг суурьтай (алдагдал -> ашиг) үед ч зөв тэмдэгтэй


def linear_forecast(series: list[float]) -> float | None:
    """y = a + b*x тааруулж дараагийн цэгийг (x=n) таамаглана."""
    n = len(series)
    if n < 3:
        return None
    xs = list(range(n))
    mx, my = statistics.mean(xs), statistics.mean(series)
    den = sum((x - mx) ** 2 for x in xs)
    slope = sum((x - mx) * (y - my) for x, y in zip(xs, series)) / den
    return my + slope * (n - mx)


def compute_bank(bd: BankData, k: tuple[int, int]) -> dict:
    y, q = k
    ann = ANNUAL[q]
    ly = (y - 1, q)
    gross, npl = bd.gross_loans_npl(k)
    prov = bd.b(k, "provision")
    deposits = bd.deposits(k)
    profit = bd.ytd(k, "net_profit")
    profit_ly = bd.ytd(ly, "net_profit")
    nii = bd.ytd(k, "nii")
    if nii is None:
        ii, ie = bd.ytd(k, "int_income"), bd.ytd(k, "int_expense")
        nii = ii - ie if None not in (ii, ie) else None
    nonint_inc, nonint_exp = bd.ytd(k, "nonint_income"), bd.ytd(k, "nonint_expense")
    assets, equity = bd.b(k, "assets"), bd.b(k, "equity")

    # Улирлын дан ашгийн прогноз: өмнөх 5 улирлын дан ашиг дээр шугаман чиг хандлага
    hist, kk = [], prev_q(y, q)
    for _ in range(5):
        hist.insert(0, bd.quarter_alone(kk, "net_profit"))
        kk = prev_q(*kk)
    hist = [h for h in hist if h is not None]
    fc = linear_forecast(hist) if len(hist) >= 4 else None
    q_profit = bd.quarter_alone(k, "net_profit")

    m = {
        "bank": bd.bank, "date": qlabel(y, q), "year": y, "quarter": q,
        "assets": assets, "assets_yoy": growth(assets, bd.b(ly, "assets")),
        "assets_qoq": growth(assets, bd.b(prev_q(y, q), "assets")),
        "loans_net": bd.b(k, "loans_net"),
        "loans_yoy": growth(bd.b(k, "loans_net"), bd.b(ly, "loans_net")),
        "deposits": deposits, "deposits_yoy": growth(deposits, bd.deposits(ly)),
        "equity": equity, "liabilities": bd.b(k, "liabilities"),
        "profit_ytd": profit, "profit_ytd_yoy": growth(profit, profit_ly),
        "profit_q": q_profit, "profit_q_prev": bd.quarter_alone(prev_q(y, q), "net_profit"),
        "profit_ttm": bd.ttm(k, "net_profit"),
        "roe": pct(profit * ann, equity) if profit is not None else None,
        "roa": pct(profit * ann, assets) if profit is not None else None,
        "nim": pct(nii * ann, assets) if nii is not None else None,
        "cost_income": pct(nonint_exp, (nii or 0) + (nonint_inc or 0)) if None not in (nonint_exp, nii, nonint_inc) else None,
        "provision_exp_ytd": bd.ytd(k, "provision_exp"),
        "npl": npl, "gross_loans": gross, "npl_ratio": pct(npl, gross),
        "coverage": pct(abs(prov), npl) if prov is not None else None,
        "ltd": pct(bd.b(k, "loans_net"), deposits),
        "equity_assets": pct(equity, assets),
        "forecast_q": fc, "forecast_gap": growth(q_profit, fc) if fc and fc > 0 else None,
        "balance_check": (assets - bd.b(k, "liabilities") - equity)
                          if None not in (assets, bd.b(k, "liabilities"), equity) else None,
    }
    return m


# ── Зах зээлийн үнэлгээ ───────────────────────────────────────────────────
def price_on_or_before(prices: dict, bank: str, iso: str):
    p = prices.get(bank)
    if not p:
        return None, None
    best = None
    for d, c in zip(p["date"], p["close"]):
        if d <= iso:
            best = (d, c)
    return best if best else (None, None)


def valuation(metrics: dict, prices: dict, shares: dict, ownership: dict | None = None) -> dict:
    """ownership (ҮЦТХТ) байвал гаргасан хувьцааны тоог ТЭНДЭЭС авна (албан ёсны төвлөрсөн
    хадгаламжийн тоо); fundamentals.json-ийнхтой зөрвөл 'shares_note'-д тэмдэглэнэ."""
    ownership = ownership or {}
    out = {}
    for bank in MSE_BANKS:
        m = metrics.get(bank)
        sh = shares.get(bank)
        own = ownership.get(bank) or {}
        note = None
        if own.get("total_shares"):
            if sh and abs(own["total_shares"] - sh) > 0.5:
                note = f"хувьцааны тоо {sh:,.0f} → {own['total_shares']:,.0f} (ҮЦТХТ {own['as_of']})"
            sh = own["total_shares"]
        if not m or not sh:
            continue
        ttm, eq = m["profit_ttm"], m["equity"]
        eps = ttm * 1e6 / sh if ttm is not None else None
        bvps = eq * 1e6 / sh if eq is not None else None
        last_d, last_p = (prices[bank]["date"][-1], prices[bank]["close"][-1]) if prices.get(bank) else (None, None)
        qend_iso = m["date"].replace("/", "-")
        qd, qp = price_on_or_before(prices, bank, qend_iso)
        out[bank] = {
            "eps": round(eps, 1) if eps else None, "bvps": round(bvps, 1) if bvps else None,
            "pe": round(last_p / eps, 2) if (last_p and eps and eps > 0) else None,
            "pb": round(last_p / bvps, 2) if (last_p and bvps) else None,
            "shares": sh, "ttm_profit": round(ttm) if ttm is not None else None,
            "equity": round(eq) if eq is not None else None,
            "price": last_p, "price_date": last_d, "price_qend": qp, "price_qend_date": qd,
            "mcap": last_p * sh / 1e6 if last_p else None,
            "as_of": m["date"], "shares_note": note,
        }
        if own:
            out[bank].update({
                "ff_pct": own.get("ff_pct"), "dom_pct": own.get("dom_pct"), "dom_holders": own.get("dom_holders"),
                "over5": own.get("over5"), "under5": own.get("under5"), "own_as_of": own.get("as_of"),
                # 5%-иас доош (чөлөөт эргэлт) ба дээш бүлгийн зах зээлийн үнэлгээ, сая ₮
                "ff_mcap": last_p * sh * own["ff_pct"] / 1e6 if (last_p and own.get("ff_pct") is not None) else None,
                "dom_mcap": last_p * sh * own["dom_pct"] / 1e6 if (last_p and own.get("dom_pct") is not None) else None,
            })
    return out


# ── Дүгнэлт (дүрэмд суурилсан) ────────────────────────────────────────────
def f0(x):
    return "—" if x is None else f"{x:,.0f}"


def fp(x, sign=False):
    if x is None:
        return "—"
    return f"{x * 100:+.1f}%" if sign else f"{x * 100:.1f}%"


def ownership_highlights(vals: dict, ownership: dict, period) -> list[str]:
    """Эзэмшлийн бүтцийн дүгнэлт: өмнөх улирлын эцсээс хойшх өөрчлөлтүүд + эзэмшигчдийн огцом өсөлт."""
    if not ownership:
        return []
    y, q = period
    since = qlabel(*prev_q(y, q)).replace("/", "-")
    lines = []
    for b in MSE_BANKS:
        for e in (ownership.get(b) or {}).get("events", []):
            if e["date"] > since:
                lines.append(f"Эзэмшлийн өөрчлөлт — {b} ({e['date']}): {e['text']}.")
    u = [(b, v["under5"]) for b, v in vals.items() if v.get("under5") and v["under5"].get("holders_3m") is not None]
    if u:
        top = sorted(u, key=lambda t: -t[1]["holders_3m"])[:2]
        lines.append("5%-иас ДООШ эзэмшигчдийн тоо 3 сард хамгийн хурдан өссөн: "
                     + ", ".join(f"{b} ({x['holders_3m']:+.1%}, {x['holders']:,.0f} эзэмшигч)" for b, x in top) + ".")
    o = [(b, v["over5"]) for b, v in vals.items() if v.get("over5")]
    if o:
        lines.append("5%-иас ДЭЭШ эзэмшигчид (тоо / эзлэх хувь): "
                     + ", ".join(f"{b} {x['holders']:.0f} / {x['pct']:.1%}" for b, x in sorted(o, key=lambda t: -t[1]['pct'])) + ".")
    ff = [(b, v["ff_pct"], v.get("ff_mcap")) for b, v in vals.items() if v.get("ff_pct") is not None]
    if ff:
        b, pct_, cap = min(ff, key=lambda t: t[1])
        lines.append(f"5%-иас ДООШ эзэмшигчдийн хувь (чөлөөт эргэлт) хамгийн бага: {b} ({pct_:.1%}"
                     + (f", ~{cap / 1000:,.1f} тэрбум ₮" if cap else "") + ") — арилжааны хөрвөх чадвар хязгаарлагдмал.")
    for b, v in vals.items():
        if v.get("shares_note"):
            lines.append(f"⚠ {b}: {v['shares_note']} — EPS/BVPS шинэ тоогоор тооцогдов.")
    return lines


def highlights(metrics: dict, period: tuple[int, int]) -> list[str]:
    cur = {b: m for b, m in metrics.items() if (m["year"], m["quarter"]) == period}
    lines = []
    tot = sum(m["assets"] for m in cur.values() if m["assets"])
    prof = sum(m["profit_ytd"] for m in cur.values() if m["profit_ytd"] is not None)
    y, q = period
    lines.append(f"{qlabel(y, q)}-ны байдлаар тайлангаа гаргасан {len(cur)} банкны нийт хөрөнгө "
                 f"{f0(tot)} сая ₮ ({tot / 1e6:.1f} их наяд ₮), {PERIOD_SHORT[q]} цэвэр ашгийн нийлбэр {f0(prof)} сая ₮.")

    def rank(key, n=3, rev=True, flt=lambda v: True):
        xs = [(b, m[key]) for b, m in cur.items() if m.get(key) is not None and flt(m[key])]
        return sorted(xs, key=lambda t: t[1], reverse=rev)[:n]

    def joined(xs, f=fp, sign=False):
        return ", ".join(f"{b} ({f(v, sign) if f is fp else f(v)})" for b, v in xs)

    if r := rank("assets_yoy"):
        lines.append("Хөрөнгийн жилийн өсөлтөөр тэргүүлэгчид: " + joined(r, sign=True) + ".")
    if r := rank("profit_ytd_yoy", flt=lambda v: v > 0):
        lines.append("Ашгийн жилийн өсөлтөөр тэргүүлэгчид: " + joined(r, sign=True) + ".")
    if r := rank("profit_ytd_yoy", n=9, rev=False, flt=lambda v: v < 0):
        lines.append("Ашиг нь өмнөх оны мөн үеэс буурсан банкууд: " + joined(r, sign=True) + ".")
    if r := rank("roe"):
        lines.append("ROE-ээр (жилийн хэмжээст) тэргүүлэгчид: " + joined(r) + ".")
    if r := rank("npl_ratio", n=9, flt=lambda v: v >= 0.06):
        lines.append("Чанаргүй зээлийн харьцаа 6%-иас дээш: " + joined(r) + ".")
    if r := rank("coverage", n=9, rev=False, flt=lambda v: v < 0.6):
        lines.append("Эрсдэлийн сангийн хучилт 60%-иас доош: " + joined(r) + ".")
    lag = [b for b, m in metrics.items() if (m["year"], m["quarter"]) < period]
    if lag:
        lines.append("Энэ улирлын тайлан хараахан ороогүй: " +
                     ", ".join(f"{b} ({metrics[b]['date']})" for b in lag) + ".")
    bad = [b for b, m in cur.items() if m["balance_check"] is not None
           and abs(m["balance_check"]) > max(1.0, 0.0005 * (m["assets"] or 0))]
    if bad:
        lines.append("⚠ Баланс тэнцэхгүй (Хөрөнгө ≠ Өр + ӨХ): " + ", ".join(bad) + " — эх файлыг шалгана уу.")
    return lines


# ── Графикууд ─────────────────────────────────────────────────────────────
def make_charts(all_data: dict[str, BankData], metrics: dict, out_dir: Path, valuation_data: dict | None = None) -> list[Path]:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams["font.family"] = "DejaVu Sans"
    paths = []

    # 1. Улирлын дан ашиг (сүүлийн 8 улирал) — МХБ-ийн 5 банк
    fig, ax = plt.subplots(figsize=(9, 4.2))
    for bank in MSE_BANKS:
        bd = all_data[bank]
        k = bd.latest()
        if not k:
            continue
        ks, kk = [], k
        for _ in range(8):
            ks.insert(0, kk)
            kk = prev_q(*kk)
        vals = [bd.quarter_alone(x, "net_profit") for x in ks]
        ax.plot([f"{a}Q{b}" for a, b in ks], [v / 1000 if v is not None else None for v in vals],
                marker="o", label=bank)
    ax.set_title("Улирлын дан цэвэр ашиг (тэрбум ₮)")
    ax.grid(alpha=.3)
    ax.legend(fontsize=8, ncol=5, loc="upper center", bbox_to_anchor=(.5, -.12))
    fig.tight_layout()
    p = out_dir / "chart_profit_quarterly.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    paths.append(p)

    # 2. ROE vs NPL (бүх банк)
    fig, ax = plt.subplots(figsize=(7, 4.5))
    for b, m in metrics.items():
        if m["roe"] is None or m["npl_ratio"] is None:
            continue
        ax.scatter(m["npl_ratio"] * 100, m["roe"] * 100, s=max(30, (m["assets"] or 0) / 1e5))
        ax.annotate(b, (m["npl_ratio"] * 100, m["roe"] * 100), fontsize=8, xytext=(4, 4), textcoords="offset points")
    ax.set_xlabel("Чанаргүй зээлийн харьцаа, %")
    ax.set_ylabel("ROE (жилийн хэмжээст), %")
    ax.set_title("Өгөөж ба эрсдэл (бөмбөлгийн хэмжээ = нийт хөрөнгө)")
    ax.grid(alpha=.3)
    fig.tight_layout()
    p = out_dir / "chart_roe_npl.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    paths.append(p)

    # 3. Нийт хөрөнгийн хэмжээ
    fig, ax = plt.subplots(figsize=(8, 4))
    items = sorted(((b, m["assets"]) for b, m in metrics.items() if m["assets"]), key=lambda t: -t[1])
    ax.barh([b for b, _ in items][::-1], [v / 1e6 for _, v in items][::-1], color="#4f7cba")
    ax.set_xlabel("их наяд ₮")
    ax.set_title("Нийт хөрөнгө (банк бүрийн сүүлийн тайлангаар)")
    ax.grid(axis="x", alpha=.3)
    fig.tight_layout()
    p = out_dir / "chart_assets.png"
    fig.savefig(p, dpi=150)
    plt.close(fig)
    paths.append(p)

    # 4. P/B ба ROE, бөмбөлөг = чөлөөт эргэлтийн үнэлгээ (ҮЦТХТ байвал)
    pts = [(b, v) for b, v in (valuation_data or {}).items() if v.get("pb") and metrics.get(b, {}).get("roe") is not None]
    if pts:
        fig, ax = plt.subplots(figsize=(7, 4.5))
        for i, (b, v) in enumerate(sorted(pts, key=lambda t: metrics[t[0]]["roe"])):
            roe = metrics[b]["roe"] * 100
            size = max(40, (v.get("ff_mcap") or v.get("mcap") or 0) / 600)
            ax.scatter(roe, v["pb"], s=size, alpha=.6)
            lab = b + (f"\nчөлөөт {v['ff_pct']:.1%}" if v.get("ff_pct") is not None else "")
            ax.annotate(lab, (roe, v["pb"]), fontsize=8, textcoords="offset points",
                        xytext=(8, 6) if i % 2 == 0 else (8, -22))   # хөрш шошго давхцахгүй
        ax.margins(x=.12, y=.18)
        ax.axhline(1, color="grey", lw=.8, ls="--")
        ax.set_xlabel("ROE (жилийн хэмжээст), %")
        ax.set_ylabel("P/B")
        ax.set_title("Үнэлгээ ба өгөөж (бөмбөлөг = чөлөөт эргэлтийн зах зээлийн үнэлгээ)")
        ax.grid(alpha=.3)
        fig.tight_layout()
        p = out_dir / "chart_pb_roe_freefloat.png"
        fig.savefig(p, dpi=150)
        plt.close(fig)
        paths.append(p)
    return paths


# ── Word тайлан ───────────────────────────────────────────────────────────
def build_docx(path: Path, period, metrics, vals, hl, charts, source_note: str,
               ownership: dict | None = None, own_note: str = "", extra_sections: list | None = None,
               own_charts: list | None = None):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    y, q = period
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2))
    sec.top_margin = sec.bottom_margin = Cm(2)
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for name, size in (("Heading 1", 14), ("Heading 2", 12), ("Title", 20)):
        s = doc.styles[name]
        s.font.name = "Arial"
        s.font.size = Pt(size)
        s.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)

    def shade(cell, hexcolor):
        tcPr = cell._tc.get_or_add_tcPr()
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear")
        sh.set(qn("w:color"), "auto")
        sh.set(qn("w:fill"), hexcolor)
        tcPr.append(sh)

    def table(headers, rows, widths=None):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        fs = 8.5 if len(headers) <= 8 else 7.5 if len(headers) <= 11 else 6.5
        for i, h in enumerate(headers):
            c = t.rows[0].cells[i]
            c.text = h
            shade(c, "1F3A5F")
            for r in c.paragraphs[0].runs:
                r.font.bold = True
                r.font.size = Pt(fs)
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        for row in rows:
            cells = t.add_row().cells
            for i, v in enumerate(row):
                cells[i].text = str(v)
                p = cells[i].paragraphs[0]
                if i > 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.font.size = Pt(fs)
        if widths:
            for row in t.rows:
                for i, w in enumerate(widths):
                    row.cells[i].width = Cm(w)
        doc.add_paragraph()

    doc.add_paragraph("МОНГОЛЫН БАНКНЫ САЛБАР", style="Title")
    doc.add_paragraph(f"{y} ОНЫ {QUARTER_ROMAN[q]} УЛИРЛЫН ШИНЖИЛГЭЭ").runs[0].bold = True
    doc.add_paragraph("Есөн арилжааны банкны санхүүгийн үзүүлэлтийн харьцуулсан судалгаа (автоматаар үүсгэсэн)")
    doc.add_paragraph(f"Үүсгэсэн огноо: {date.today():%Y-%m-%d}    Тайлант хугацаа: {qlabel(y, q)}")
    doc.add_paragraph(source_note).runs[0].italic = True

    doc.add_heading("1. Хураангуй", level=1)
    for line in hl:
        doc.add_paragraph(line, style="List Bullet")

    doc.add_heading("2. Хэмжээ ба ашигт ажиллагаа", level=1)
    order = sorted(metrics, key=lambda b: -(metrics[b]["assets"] or 0))
    tot = sum(metrics[b]["assets"] or 0 for b in order)
    table(["Банк", "Огноо", "Хөрөнгө (сая ₮)", "YoY", "Эзлэх %", f"Ашиг YTD", "YoY", "ROE", "ROA"],
          [[b, metrics[b]["date"][2:], f0(metrics[b]["assets"]), fp(metrics[b]["assets_yoy"], True),
            fp(pct(metrics[b]["assets"], tot)), f0(metrics[b]["profit_ytd"]), fp(metrics[b]["profit_ytd_yoy"], True),
            fp(metrics[b]["roe"]), fp(metrics[b]["roa"])] for b in order],
          [2.6, 1.7, 2.3, 1.4, 1.4, 2.1, 1.5, 1.4, 1.4])

    doc.add_heading("3. Активын чанар, үр ашиг, хөрвөх чадвар", level=1)
    table(["Банк", "NPL %", "Хучилт", "Зээл/Харилцах+Хадгаламж", "ӨХ/Хөрөнгө", "NIM", "Зардал/Орлого"],
          [[b, fp(metrics[b]["npl_ratio"]), fp(metrics[b]["coverage"]), fp(metrics[b]["ltd"]),
            fp(metrics[b]["equity_assets"]), fp(metrics[b]["nim"]), fp(metrics[b]["cost_income"])] for b in order],
          [2.8, 1.8, 1.8, 3.4, 2.2, 1.8, 2.4])
    doc.add_paragraph("Тайлбар: NPL = чанаргүй зээл / нийт (бохир) зээл, нийт зээл = цэвэр зээл + эрсдэлийн сан; "
                      "хучилт = зээлийн эрсдэлийн сан / чанаргүй зээл; ROE, ROA, NIM нь тайлант үеийн хуримтлалыг "
                      "жилийн хэмжээст болгосон (Q1×4, H1×2, 9 сар×4/3, жил×1), хугацааны эцсийн үлдэгдэлд харьцуулсан.")

    if vals:
        doc.add_heading("4. Зах зээлийн үнэлгээ (МХБ-ийн 5 банк)", level=1)
        table(["Банк", "Ханш ₮", "Ханшийн огноо", "EPS (TTM)", "BVPS", "P/E", "P/B", "Зах зээлийн үнэлгээ (сая ₮)"],
              [[b, f0(v["price"]), v["price_date"] or "—", f0(v["eps"]), f0(v["bvps"]),
                f"{v['pe']:.2f}x" if v["pe"] else "—", f"{v['pb']:.2f}x" if v["pb"] else "—", f0(v["mcap"])]
               for b, v in vals.items()],
              [2.6, 1.7, 2.3, 1.8, 1.8, 1.4, 1.4, 3.2])
        doc.add_paragraph("EPS = сүүлийн 4 улирлын (TTM) цэвэр ашиг / гаргасан хувьцаа; BVPS = өөрийн хөрөнгө / гаргасан хувьцаа.")

    sec = 5
    if ownership and any(v.get("ff_pct") is not None for v in vals.values()):
        doc.add_heading("5. Эзэмшлийн бүтэц ба хувьцааны хөрвөх чадвар (МХБ-ийн 5 банк)", level=1)
        doc.add_paragraph(f"Эх сурвалж: {own_note}. Хувьцаа эзэмшигчдийг ҮЦТХТ-ийн ангиллаар хоёр тусдаа бүлэгт "
                          "харуулав: 5%-иас ДЭЭШ эзэмшигчид (хяналтын/стратегийн эзэмшигчид) ба 5%-иас ДООШ эзэмшигчид "
                          "(олон нийт, чөлөөт эргэлт). Үнэлгээ = сүүлийн ханш × бүлгийн хувьцаа.")
        ov = [(b, v) for b, v in sorted(vals.items(), key=lambda t: -(t[1].get("dom_pct") or 0)) if v.get("over5")]

        def pp(x):
            return "—" if x is None else f"{x * 100:+.2f}"

        def dn(x):
            return "—" if x is None else f"{x:+.0f}"
        doc.add_heading("5%-иас ДЭЭШ эзэмшигчид", level=2)
        table(["Банк", "Эзэмшигчийн тоо", "Тоо: 3 сар", "Тоо: 12 сар", "Эзлэх хувь", "н.х.: 3 сар", "н.х.: 12 сар",
               "Хувьцаа, ш", "Үнэлгээ, тэрбум ₮"],
              [[b, f0(v["over5"]["holders"]), dn(v["over5"]["holders_chg_3m"]), dn(v["over5"]["holders_chg_12m"]),
                f"{v['over5']['pct']:.2%}", pp(v["over5"]["pct_chg_3m"]), pp(v["over5"]["pct_chg_12m"]),
                f0(v["over5"]["shares"]), f0((v.get("dom_mcap") or 0) / 1000) if v.get("dom_mcap") else "—"] for b, v in ov],
              [2.4, 1.7, 1.5, 1.5, 1.7, 1.6, 1.6, 2.6, 2.0])
        un = [(b, v) for b, v in sorted(vals.items(), key=lambda t: -((t[1].get("under5") or {}).get("holders") or 0))
              if v.get("under5")]
        doc.add_heading("5%-иас ДООШ эзэмшигчид (олон нийт, чөлөөт эргэлт)", level=2)
        table(["Банк", "Эзэмшигчийн тоо", "3 сар", "12 сар", "Эзлэх хувь", "н.х.: 12 сар", "Хувьцаа, ш",
               "Нэг эзэмшигчид, ш", "Үнэлгээ, тэрбум ₮"],
              [[b, f0(v["under5"]["holders"]), fp(v["under5"]["holders_3m"], True), fp(v["under5"]["holders_12m"], True),
                f"{v['under5']['pct']:.2%}", pp(v["under5"]["pct_chg_12m"]), f0(v["under5"]["shares"]),
                f0(v["under5"]["avg_shares"]), f0((v.get("ff_mcap") or 0) / 1000) if v.get("ff_mcap") else "—"] for b, v in un],
              [2.4, 1.7, 1.4, 1.4, 1.6, 1.6, 2.4, 1.9, 2.0])
        rr = [[b, f"{v['dom_pct']:.1%}", fp(v["ff_pct"]), f"{v['pb']:.2f}x" if v.get("pb") else "—", fp(metrics[b]["roe"])]
              for b, v in ov]
        doc.add_heading("Эзэмшлийн бүтэц ба үнэлгээ", level=2)
        table(["Банк", "5%-иас дээш", "5%-иас доош", "P/B", "ROE"], rr, [3.2, 3.2, 3.2, 3.2, 3.2])
        for c in own_charts or []:
            doc.add_picture(str(c), width=Cm(16.5))
        evs = [(b, e) for b in MSE_BANKS for e in (ownership.get(b) or {}).get("events", [])]
        if evs:
            doc.add_paragraph("Сүүлийн 12 сарын эзэмшлийн бүтцийн өөрчлөлтүүд:")
            for b, e in sorted(evs, key=lambda t: t[1]["date"], reverse=True):
                doc.add_paragraph(f"{e['date']} — {b}: {e['text']}", style="List Bullet")
        else:
            doc.add_paragraph("Сүүлийн 12 сард эзэмшлийн бүтцэд мэдэгдэхүйц өөрчлөлт гараагүй.")
        sec = 6

    doc.add_heading(f"{sec}. Банк тус бүр", level=1)
    for b in order:
        m = metrics[b]
        doc.add_heading(b, level=2)
        yy, qq = m["year"], m["quarter"]
        txt = (f"{b}-ны нийт хөрөнгө {m['date']}-ны байдлаар {f0(m['assets'])} сая ₮ "
               f"(өмнөх оны мөн үеэс {fp(m['assets_yoy'], True)}, өмнөх улирлаас {fp(m['assets_qoq'], True)}). "
               f"Цэвэр зээл {f0(m['loans_net'])} сая ₮ ({fp(m['loans_yoy'], True)} YoY), харилцах ба хадгаламж "
               f"{f0(m['deposits'])} сая ₮ ({fp(m['deposits_yoy'], True)} YoY).")
        doc.add_paragraph(txt)
        txt = (f"{yy} оны {PERIOD_MN[qq]}-ийн цэвэр ашиг {f0(m['profit_ytd'])} сая ₮ "
               f"(өмнөх оны мөн үеэс {fp(m['profit_ytd_yoy'], True)}); зөвхөн энэ улирлын ашиг {f0(m['profit_q'])} сая ₮ "
               f"(өмнөх улиралд {f0(m['profit_q_prev'])}). ROE {fp(m['roe'])}, ROA {fp(m['roa'])}, NIM {fp(m['nim'])}.")
        doc.add_paragraph(txt)
        txt = f"Чанаргүй зээлийн харьцаа {fp(m['npl_ratio'])}, эрсдэлийн сангийн хучилт {fp(m['coverage'])}."
        if m["forecast_q"] is not None and m["profit_q"] is not None:
            txt += (f" Өмнөх 5 улирлын дан ашгийн шугаман чиг хандлагаар энэ улирлын ашиг {f0(m['forecast_q'])} сая ₮ "
                    f"гэж тооцоологдож байсан; бодит {f0(m['profit_q'])} сая ₮"
                    + (f" ({fp(m['forecast_gap'], True)})." if m["forecast_gap"] is not None else "."))
        doc.add_paragraph(txt)
        if b in vals:
            v = vals[b]
            t = (f"Ханш {f0(v['price'])}₮ ({v['price_date']}), улирлын эцэст {f0(v['price_qend'])}₮. "
                 f"P/E {v['pe'] if v['pe'] else '—'}x, P/B {v['pb'] if v['pb'] else '—'}x.")
            if v.get("over5") and v.get("under5"):
                o_, u_ = v["over5"], v["under5"]
                t += (f" Хувьцаа эзэмшигчид ({v['own_as_of']}): 5%-иас ДЭЭШ {o_['holders']:.0f} эзэмшигч нийт "
                      f"{o_['pct']:.1%}-ийг; 5%-иас ДООШ {f0(u_['holders'])} эзэмшигч {u_['pct']:.1%}-ийг эзэмшиж байна "
                      f"(тоо нь 3 сард {fp(u_['holders_3m'], True)}, 12 сард {fp(u_['holders_12m'], True)}).")
            doc.add_paragraph(t)

    nxt = sec + 1
    for spec in extra_sections or []:
        doc.add_heading(f"{nxt}. {spec['title']}", level=1)
        if spec.get("intro"):
            doc.add_paragraph(spec["intro"])
        for title, headers, rows in spec.get("tables", []):
            if title:
                doc.add_heading(title, level=2)
            n = len(headers)
            first = 2.6 if n <= 8 else 2.3
            table(headers, rows, [first] + [(16.8 - first) / (n - 1)] * (n - 1))
        for line in spec.get("bullets", []):
            doc.add_paragraph(line, style="List Bullet")
        for note in spec.get("notes", []):
            r = doc.add_paragraph().add_run(note)
            r.italic = True
            r.font.size = Pt(9)
        for c in spec.get("charts", []):
            doc.add_picture(str(c), width=Cm(16.5))
        nxt += 1

    doc.add_heading(f"{nxt}. Графикууд", level=1)
    for c in charts:
        doc.add_picture(str(c), width=Cm(16))
    doc.add_paragraph()
    doc.add_paragraph("Тайлбар: Энэ тайланг workbook-ийн өгөгдлөөс автоматаар үүсгэсэн. Тоон дүн нь эх тайлангаас шууд "
                      "гарсан боловч дүгнэлтийн өгүүлбэрүүд нь дүрэмд суурилсан тул эцсийн дүгнэлт, хөрөнгө оруулалтын "
                      "шийдвэр гаргахын өмнө хянаж үзнэ үү. Зөвхөн мэдээллийн зорилготой.").runs[0].italic = True
    doc.save(path)


def telegram_text(period, metrics, vals, hl) -> str:
    y, q = period
    lines = [f"📊 <b>{y} оны {QUARTER_ROMAN[q]} улирлын банкны шинжилгээ</b>", ""]
    lines += [f"• {h}" for h in hl]
    lines.append("")
    for b in MSE_BANKS:
        m = metrics.get(b)
        if not m:
            continue
        v = vals.get(b, {})
        lines.append(f"<b>{b}</b> ({m['date'][2:]}): ашиг {f0(m['profit_ytd'])} ({fp(m['profit_ytd_yoy'], True)}), "
                     f"ROE {fp(m['roe'])}, NPL {fp(m['npl_ratio'])}"
                     + (f", P/E {v['pe']}x, P/B {v['pb']}x" if v.get("pe") else "")
                     + (f", 5%+ : {v['over5']['holders']:.0f} эзэмшигч {v['over5']['pct']:.1%}, 5%- : "
                        f"{f0(v['under5']['holders'])} эзэмшигч {v['under5']['pct']:.1%} ({fp(v['under5']['holders_3m'], True)} 3с)"
                        if v.get("over5") and v.get("under5") else ""))
    return "\n".join(lines)


def run(workbook: str, out_root: str = "analysis", prices_path: str | None = None,
        fundamentals_path: str | None = None, source_note: str = "",
        ownership: dict | None = None, ownership_note: str = "",
        ownership_url: str | None = None, ownership_cache: str | None = None,
        ownership_file: str | None = None, adv_params: dict | None = None,
        status_lines: list[str] | None = None) -> dict:
    """ownership: None бол ownership_source.load()-оор GitHub-аас татна; {} бол алгасна."""
    wb = load_workbook(workbook, data_only=True)
    data = {b: BankData(wb, b) for b in BANKS if f"{b} баланс" in wb.sheetnames}
    metrics = {}
    for b, bd in data.items():
        k = bd.latest()
        if k:
            metrics[b] = compute_bank(bd, k)
    period = max((m["year"], m["quarter"]) for b, m in metrics.items() if b in MSE_BANKS)
    y, q = period
    out_dir = Path(out_root) / f"{y}Q{q}"
    out_dir.mkdir(parents=True, exist_ok=True)

    prices = json.loads(Path(prices_path).read_text(encoding="utf-8")) if prices_path and Path(prices_path).exists() and Path(prices_path).stat().st_size else {}
    shares = {}
    fund_old = {}
    if fundamentals_path and Path(fundamentals_path).exists():
        fund_old = json.loads(Path(fundamentals_path).read_text(encoding="utf-8"))
        shares = {b: v.get("shares") for b, v in fund_old.items()}
    if ownership is None:
        try:
            import ownership_source
            ownership, ownership_note = ownership_source.load(
                ownership_url or ownership_source.DEFAULT_URL,
                Path(ownership_cache) if ownership_cache else ownership_source.DEFAULT_CACHE,
                ownership_file, quarter_end_iso=qlabel(y, q).replace("/", "-"))
        except Exception as e:  # noqa: BLE001
            ownership, ownership_note = {}, f"эзэмшлийн мэдээлэл ороогүй: {e}"
    if ownership_note:
        print(f"  ℹ {ownership_note}")
    if ownership:
        for b, o in ownership.items():
            shares.setdefault(b, o.get("total_shares"))
    vals = valuation(metrics, prices, shares, ownership) if prices and shares else {}

    hl = highlights(metrics, period) + ownership_highlights(vals, ownership or {}, period)
    import extended_analysis as _ea
    own_charts = _ea.ownership_charts(vals, ownership or {}, out_dir) if ownership else []
    charts = make_charts(data, metrics, out_dir, vals)

    # Нэмэлт харьцаа үзүүлэлт + графикууд (advanced_ratios.py)
    import advanced_ratios as ar
    ar.attach_extra(wb, data)
    aparams = {**ar.DEFAULTS, **{k: v for k, v in (adv_params or {}).items() if v is not None}}
    adv = {b: ar.compute_advanced(data[b], (m["year"], m["quarter"]), m, aparams) for b, m in metrics.items()}
    jpb = ar.justified_pb(adv, vals, metrics, aparams)
    hhi = ar.sector_hhi(data)
    adv_charts = ar.make_advanced_charts(data, metrics, adv, jpb, hhi, out_dir, aparams)
    hl += ar.adv_highlights(adv, jpb, aparams)

    # Өргөтгөсөн шинжилгээ (extended_analysis.py)
    import extended_analysis as ea
    ext = ea.compute(data, metrics, adv, vals, prices, ownership or {})
    ext_charts = ea.make_charts(data, metrics, adv, ext, out_dir)
    hl += ea.highlights(ext, metrics)
    order_ = sorted(metrics, key=lambda b: -(metrics[b]["assets"] or 0))
    adv_section = {"title": "Нэмэлт харьцаа: эрсдэл, эх үүсвэр, DuPont, капитал, үнэлгээ",
                   "intro": f"Таамаглал: COE {aparams['coe']:.0%}, урт хугацааны өсөлт g {aparams['growth']:.0%}, стресс тестийн "
                            f"LGD {aparams['lgd']:.0%}. '—' = тухайн банкны хуудсанд шаардлагатай мөр алга (тэг гэж үзээгүй).",
                   "tables": ar.table_rows(adv, order_), "bullets": ar.DEFINITIONS, "charts": adv_charts}
    import re as _re
    latest_price = max((p_["date"][-1] for p_ in prices.values()), default="—") if prices else "—"
    status = [_re.sub(r"<[^>]+>", "", l) for l in (status_lines or [])]
    status += [f"{b}: сүүлийн тайлан {metrics[b]['date']}" for b in order_]
    status += [f"Ханшийн өгөгдөл {latest_price} хүртэл (МХБ).", f"Эзэмшлийн мэдээлэл: {ownership_note or '—'}.",
               f"Workbook: {Path(workbook).name}."]
    sections = [adv_section] + ea.sections(ext, metrics, ext_charts, vals) + [
        {"title": "Өгөгдлийн төлөв ба эх сурвалж", "bullets": status}]
    own_out = {b: {k: v for k, v in o.items() if k != "history"} for b, o in (ownership or {}).items()}
    (out_dir / "metrics.json").write_text(json.dumps({"period": f"{y}Q{q}", "metrics": metrics, "valuation": vals,
                                                       "ownership": own_out, "ownership_source": ownership_note,
                                                       "advanced": adv, "justified_pb": jpb, "sector_hhi": hhi,
                                                       "extended": {k: v for k, v in ext.items() if k != "_vals"},
                                                       "advanced_params": aparams,
                                                       "highlights": hl}, ensure_ascii=False, indent=1), encoding="utf-8")
    keys = list(next(iter(metrics.values())).keys())
    with open(out_dir / "metrics.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        w.writerows(metrics.values())
    build_docx(out_dir / f"Банкны_салбар_{y}Q{q}.docx", period, metrics, vals, hl, charts,
               source_note or f"Эх сурвалж: {Path(workbook).name} (банк тус бүрийн улирлын тайлан), МХБ-ийн арилжааны түүх"
               + (", ҮЦТХТ-ийн төвлөрлийн судалгаа." if ownership else "."), ownership, ownership_note,
               sections, own_charts)
    tg = telegram_text(period, metrics, vals, hl)
    (out_dir / "telegram.txt").write_text(tg, encoding="utf-8")

    if fundamentals_path and vals:
        # dashboard-ийн хүлээж буй түлхүүрүүдийг хадгалж, шинэ утгаар солино
        new_f = {}
        for b, v in vals.items():
            if v["eps"] is None:
                new_f[b] = fund_old.get(b, {})
                continue
            new_f[b] = {"eps": v["eps"], "bvps": v["bvps"], "pe": v["pe"] or 0, "pb": v["pb"] or 0,
                        "shares": v["shares"], "ttm_profit": v["ttm_profit"], "equity": v["equity"],
                        "as_of": v["as_of"]}
            if v.get("ff_pct") is not None:
                new_f[b].update({"free_float": v["ff_pct"],
                                 "holders_over5": (v.get("over5") or {}).get("holders"), "pct_over5": v.get("dom_pct"),
                                 "holders_under5": (v.get("under5") or {}).get("holders"), "pct_under5": v["ff_pct"],
                                 "ownership_as_of": v["own_as_of"]})
        Path(fundamentals_path).write_text(json.dumps(new_f, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"✓ Шинжилгээ: {out_dir}  ({len(metrics)} банк, үндсэн хугацаа {qlabel(y, q)})")
    return {"period": period, "out_dir": out_dir, "metrics": metrics, "valuation": vals,
            "highlights": hl, "telegram": tg, "advanced": adv, "extended": ext,
            "charts": charts + adv_charts + [c for v in ext_charts.values() for c in v],
            "docx": out_dir / f"Банкны_салбар_{y}Q{q}.docx"}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--out", default="analysis")
    ap.add_argument("--prices", default=None)
    ap.add_argument("--fundamentals", default=None)
    ap.add_argument("--ownership-file", default=None, help="локал uctht_concentration.xlsx (татахгүй)")
    ap.add_argument("--no-ownership", action="store_true")
    ap.add_argument("--coe", type=float, default=None, help="хувьцааны өгөөжийн шаардлага (жиш. 0.20)")
    ap.add_argument("--growth", type=float, default=None, help="урт хугацааны өсөлт g (жиш. 0.06)")
    ap.add_argument("--lgd", type=float, default=None, help="стресс тестийн алдагдлын хувь (жиш. 0.5)")
    a = ap.parse_args()
    own, note = None, ""
    if a.no_ownership:
        own = {}
    elif a.ownership_file:
        import ownership_source
        own, note = ownership_source.load(file=a.ownership_file)
    run(a.workbook, a.out, a.prices, a.fundamentals, ownership=own, ownership_note=note,
        adv_params={"coe": a.coe, "growth": a.growth, "lgd": a.lgd})
