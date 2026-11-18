#!/usr/bin/env python3
"""
extended_analysis.py — банкны шинжилгээний ӨРГӨТГӨСӨН багц (bank_analysis.run дуудна).

Workbook-ийн түүхий баланс/орлого, МХБ-ийн ханш/арилжааны хэмжээ, ҮЦТХТ-ийн эзэмшлийн
өгөгдлөөс доорх бүлгүүдийг тооцож, хүснэгт + график + дүгнэлт болгон буцаана:

  A. Өсөлт ба зах зээлийн хувь     YoY/QoQ өсөлт (хөрөнгө, зээл, хадгаламж, ӨХ, ЦХО, зардал,
                                   ашиг), 2 жилийн CAGR, зах зээлийн хувь ба өөрчлөлт (н.х.)
  B. Балансын ба орлогын бүтэц     Активын бүтэц (мөнгө, байршуулсан, хөрөнгө оруулалт, зээл,
                                   бусад), зээл/хөрөнгө, хүүгийн бус орлогын хувь
  C. Өгөөж, өртөг, маржин          Хүү өгөх активын өгөөж, хүү төлөх эх үүсвэрийн өртөг, спред,
                                   NIM (хүү өгөх активт)
  D. Ашгийн чанар ба үр ашиг       Сан/ЦХО, сан/PPOP, татварын бодит хувь, зардал/хөрөнгө,
                                   орлого/хөрөнгө, татварын өмнөх ROA, TTM ROE, улирлын ROE,
                                   улирлын ашгийн хэлбэлзэл (CV)
  E. Өргөтгөсөн эрсдэл             Нөөцийн харьцаа, NPL/ӨХ, цэвэр NPL/ӨХ, мөнгө/хадгаламж,
                                   зээл/нийт эх үүсвэр, бөөний эх үүсвэр, ӨХ/зээл
  F. Нэгдсэн скоркард              11 үзүүлэлтийн эрэмбэ → нийлмэл оноо (0–100)
  G. Хувьцааны гүйцэтгэл           1с/3с/6с/1ж/YTD өгөөж, хэлбэлзэл, хамгийн их уналт,
                                   52 долоо хоногийн оргилоос, өдрийн дундаж арилжаа, эргэлт
  H. Өргөтгөсөн үнэлгээ            Ашгийн өгөөж (E/P), ROE/PB, зах зээлийн үнэлгээ/хөрөнгө,
                                   Graham тоо ба зөрүү, PEG
  I. Прогноз                       Дараагийн улирлын дан ашиг (шугаман чиг хандлага ± хазайлт),
                                   жилийн ашгийн тооцоо
  J. Чиг хандлага                  NIM, зардал/орлого, ӨХ/хөрөнгө, зээл/хадгаламж — 10 улирал

Бүх харьцаа мэдээлэл дутуу бол None ("—") — тэг гэж үзэхгүй.
"""
from __future__ import annotations

import math
import statistics
from datetime import date, timedelta
from pathlib import Path

import advanced_ratios as ar
import bank_analysis as ba

fp, f0 = ba.fp, ba.f0


def _sd(a, b):
    return None if a is None or b in (None, 0) else a / b


def _g(a, b):
    return ba.growth(a, b)


def _sum(*xs):
    xs = [x for x in xs if x is not None]
    return sum(xs) if xs else None


def _avg(a, b):
    return a if b is None else (b if a is None else (a + b) / 2)


# ═════════════════════════ ТООЦОО ══════════════════════════════════════════
def bank_block(bd: ba.BankData, k, m: dict) -> dict:
    y, q = k
    ann = ba.ANNUAL[q]
    ly, pq, ys = (y - 1, q), ba.prev_q(y, q), (y - 1, 4)
    b, x = bd.b, lambda kk, key: ar._extra(bd, kk, key)
    A, E, L = m["assets"], m["equity"], m["liabilities"]
    gross, npl = bd.gross_loans_npl(k)
    prov = abs(b(k, "provision")) if b(k, "provision") is not None else None
    dep = m["deposits"]

    def yt(key, kk=k):
        return bd.ytd(kk, key)

    nii, ninc, nexp, pexp = yt("nii"), yt("nonint_income"), yt("nonint_expense"), yt("provision_exp")
    ii, ie, pbt, np_ = yt("int_income"), yt("int_expense"), yt("pbt"), yt("net_profit")

    def iea(kk):
        """Хүү өгөх актив. Байршуулсан/хөрөнгө оруулалтын мөр хуудсанд БАЙХГҮЙ бол None —
        дутуу хуваагчаар өгөөжийг хэтрүүлэхгүйн тулд (жиш. Капитрон)."""
        parts = [x(kk, "placements"), x(kk, "investments"), b(kk, "loans_net")]
        return sum(parts) if all(p is not None for p in parts) else None

    def ibl(kk):
        d = bd.deposits(kk)
        return _sum(d, x(kk, "bank_borrow"), x(kk, "other_src")) if d is not None else None

    avgA = _avg(A, b(ys, "assets"))
    avg_iea, avg_ibl = _avg(iea(k), iea(ys)), _avg(ibl(k), ibl(ys))
    ppop = (nii + ninc - nexp) if None not in (nii, ninc, nexp) else None

    # улирлын дан ашгийн хэлбэлзэл (сүүлийн 8 улирал)
    qs, kk = [], k
    for _ in range(8):
        qs.append(bd.quarter_alone(kk, "net_profit"))
        kk = ba.prev_q(*kk)
    qs = [v for v in qs if v is not None]
    cv = (statistics.pstdev(qs) / abs(statistics.mean(qs))) if len(qs) >= 4 and statistics.mean(qs) else None

    cash, plc, inv, loans = b(k, "cash"), x(k, "placements"), x(k, "investments"), b(k, "loans_net")
    other_assets = (A - _sum(cash, plc, inv, loans)) if A and loans is not None else None
    funding = ibl(k)
    a2 = b((y - 2, q), "assets")

    return {
        # A. өсөлт
        "g_assets": m["assets_yoy"], "g_loans": m["loans_yoy"], "g_deposits": m["deposits_yoy"],
        "g_equity": _g(E, b(ly, "equity")), "g_nii": _g(nii, yt("nii", ly)), "g_opex": _g(nexp, yt("nonint_expense", ly)),
        "g_profit": m["profit_ytd_yoy"], "g_int_income": _g(ii, yt("int_income", ly)),
        "q_assets": m["assets_qoq"], "q_loans": _g(loans, b(pq, "loans_net")), "q_deposits": _g(dep, bd.deposits(pq)),
        "cagr2_assets": ((A / a2) ** .5 - 1) if A and a2 else None,
        # B. бүтэц
        "mix_cash": _sd(cash, A), "mix_placements": _sd(plc, A), "mix_investments": _sd(inv, A),
        "mix_loans": _sd(loans, A), "mix_other": _sd(other_assets, A),
        "deposits_assets": _sd(dep, A), "nonint_share": _sd(ninc, (nii or 0) + (ninc or 0)) if None not in (nii, ninc) else None,
        # C. өгөөж, өртөг
        "asset_yield": _sd(ii * ann, avg_iea) if ii is not None else None,
        "funding_cost": _sd(ie * ann, avg_ibl) if ie is not None else None,
        "nim_iea": _sd(nii * ann, avg_iea) if nii is not None else None,
        # D. ашгийн чанар
        "prov_nii": _sd(pexp, nii), "prov_ppop": _sd(pexp, ppop),
        # ТӨА ≈ цэвэр ашиг (зөрүү <0.5%) бол татварын мөр дутуу гэж үзнэ (жиш. ХХБ) — 0% биш "—"
        "eff_tax": _sd(pbt - np_, pbt) if None not in (pbt, np_) and pbt > 0 and (pbt - np_) > .005 * pbt else None,
        "opex_assets": _sd(nexp * ann, avgA) if nexp is not None else None,
        "rev_assets": _sd((nii + ninc) * ann, avgA) if None not in (nii, ninc) else None,
        "roa_pretax": _sd(pbt * ann, avgA) if pbt is not None else None,
        "roe_ttm": _sd(m["profit_ttm"], _avg(E, b(ly, "equity"))),
        "roe_quarter": _sd(m["profit_q"] * 4, E) if m["profit_q"] is not None else None,
        "profit_cv": cv,
        # E. өргөтгөсөн эрсдэл
        "reserve_ratio": _sd(prov, gross), "npl_equity": _sd(npl, E),
        "net_npl_equity": _sd(npl - prov, E) if None not in (npl, prov) else None,
        "cash_deposits": _sd(cash, dep), "loans_funding": _sd(loans, funding),
        "wholesale_share": _sd(_sum(x(k, "bank_borrow"), x(k, "other_src")), L),
        "equity_loans": _sd(E, loans),
    }


def market_share(data: dict) -> dict:
    """9 банк бүгд өгөгдөлтэй хамгийн сүүлийн улиралд зах зээлийн хувь ба жилийн өмнөхөөс өөрчлөлт."""
    common = sorted(set.intersection(*[set(bd.bal) for bd in data.values()]))
    if not common:
        return {}
    k = common[-1]
    ly = (k[0] - 1, k[1])
    out = {"period": f"{k[0]}Q{k[1]}", "banks": {}}
    for key, fn in (("assets", lambda bd, kk: bd.b(kk, "assets")), ("loans", lambda bd, kk: bd.b(kk, "loans_net")),
                    ("deposits", lambda bd, kk: bd.deposits(kk))):
        now = {b: fn(bd, k) for b, bd in data.items()}
        prev = {b: fn(bd, ly) for b, bd in data.items()} if ly in common else {}
        tn = sum(v for v in now.values() if v) or None
        tp = sum(v for v in prev.values() if v) or None
        for b in data:
            s = _sd(now[b], tn)
            sp = _sd(prev.get(b), tp)
            out["banks"].setdefault(b, {})[key] = s
            out["banks"][b][key + "_chg"] = (s - sp) if None not in (s, sp) else None
    return out


def stock_stats(prices: dict, vals: dict, ownership: dict) -> dict:
    out = {}
    for b, p in prices.items():
        if b not in vals:
            continue
        d = [date.fromisoformat(x) for x in p["date"]]
        c, vol = p["close"], p.get("volume") or [0] * len(c)
        last = d[-1]

        def back(days):
            t = last - timedelta(days=days)
            i = max((j for j, dd in enumerate(d) if dd <= t), default=None)
            return c[i] if i is not None else None
        ytd0 = next((c[j] for j, dd in enumerate(d) if dd.year == last.year), None)
        yr = [j for j, dd in enumerate(d) if dd > last - timedelta(days=365)]
        rets = [math.log(c[j] / c[j - 1]) for j in yr if j > 0 and c[j - 1] and c[j]]
        peak, mdd = 0, 0.0
        for j in yr:
            peak = max(peak, c[j])
            mdd = min(mdd, c[j] / peak - 1)
        hi52 = max(c[j] for j in yr) if yr else None
        q3 = [j for j, dd in enumerate(d) if dd > last - timedelta(days=91)]
        traded = [c[j] * (vol[j] or 0) for j in q3]
        days_traded = sum(1 for j in q3 if (vol[j] or 0) > 0)
        # Ханшийн өгөгдөлд зөвхөн арилжаатай өдрүүд байдаг тул хуваагч = тухайн 91 хоногийн АЖЛЫН өдөр
        bdays = sum(1 for i in range(91) if (last - timedelta(days=i)).weekday() < 5)
        sh = vals[b].get("shares")
        ff = (ownership.get(b) or {}).get("ff_pct")
        vol_1y = sum((vol[j] or 0) for j in yr)
        out[b] = {
            "last_date": last.isoformat(), "price": c[-1],
            "r_1m": _g(c[-1], back(30)), "r_3m": _g(c[-1], back(91)), "r_6m": _g(c[-1], back(182)),
            "r_1y": _g(c[-1], back(365)), "r_ytd": _g(c[-1], ytd0),
            "vol_1y": statistics.pstdev(rets) * math.sqrt(len(rets)) if len(rets) > 20 else None,
            "max_dd_1y": mdd if yr else None, "from_high": _g(c[-1], hi52),
            "adtv_3m": (sum(traded) / bdays / 1e6) if q3 else None,          # сая ₮ / ажлын өдөр
            "days_traded_3m": min(1.0, _sd(days_traded, bdays)) if bdays else None,
            "turnover_1y": _sd(vol_1y, sh * ff if (sh and ff) else sh),
        }
    return out


def valuation_extra(vals: dict, metrics: dict) -> dict:
    out = {}
    for b, v in vals.items():
        m = metrics[b]
        eps, bvps, p = v.get("eps"), v.get("bvps"), v.get("price")
        graham = math.sqrt(22.5 * eps * bvps) if eps and bvps and eps > 0 and bvps > 0 else None
        g = m.get("profit_ytd_yoy")
        out[b] = {
            "earnings_yield": _sd(eps, p), "roe_pb": _sd(m["roe"], v.get("pb")),
            "mcap_assets": _sd(v.get("mcap"), m["assets"]),
            "graham": graham, "graham_gap": _g(graham, p) if graham else None,
            "peg": (v["pe"] / (g * 100)) if v.get("pe") and g and g > 0 else None,
        }
    return out


def forecast(bd: ba.BankData, k, n_hist: int = 6) -> dict:
    """Дараагийн улирлын дан ашиг: сүүлийн n_hist улирлын шугаман чиг хандлага;
    муж = ± үлдэгдлийн стандарт хазайлт. q<4 бол жилийн ашгийн тооцоо = YTD + үлдсэн улирлууд."""
    hist, kk = [], k
    for _ in range(n_hist):
        hist.insert(0, (kk, bd.quarter_alone(kk, "net_profit")))
        kk = ba.prev_q(*kk)
    ys = [v for _, v in hist if v is not None]
    if len(ys) < 4:
        return {}
    n = len(ys)
    xs = list(range(n))
    mx, my = statistics.mean(xs), statistics.mean(ys)
    slope = sum((a - mx) * (b - my) for a, b in zip(xs, ys)) / sum((a - mx) ** 2 for a in xs)
    icpt = my - slope * mx
    resid = statistics.pstdev([b - (icpt + slope * a) for a, b in zip(xs, ys)])
    y, q = k
    nxt = (y, q + 1) if q < 4 else (y + 1, 1)
    f1 = icpt + slope * n
    fy = None
    if q < 4:
        ytd = bd.ytd(k, "net_profit")
        rest = [icpt + slope * (n + i) for i in range(4 - q)]
        fy = ytd + sum(rest) if ytd is not None else None
    return {"next": f"{nxt[0]}Q{nxt[1]}", "forecast": f1, "low": f1 - resid, "high": f1 + resid,
            "trend_per_q": slope, "fy_estimate": fy, "fy_year": y if q < 4 else None,
            "history": [(f"{a}Q{c}", v) for (a, c), v in hist]}


SCORE_METRICS = [  # (нэр, эх, түлхүүр, илүү сайн нь)
    ("ӨХ/хөрөнгө", "adv", "stress_base", +1), ("NPL", "m", "npl_ratio", -1), ("Хучилт", "m", "coverage", +1),
    ("Texas", "adv", "texas", -1), ("ROA", "m", "roa", +1), ("ROE", "m", "roe", +1), ("NIM", "m", "nim", +1),
    ("Зардал/орлого", "m", "cost_income", -1), ("Хөрвөх хөр./хөр.", "adv", "liquid_assets", +1),
    ("Хадг./өр", "adv", "deposit_share", +1), ("PPOP/хөр.", "adv", "ppop_assets", +1),
]


def scorecard(metrics, adv) -> dict:
    banks = list(metrics)
    ranks = {b: {} for b in banks}
    for name, src, key, sign in SCORE_METRICS:
        vals = [(b, (adv[b] if src == "adv" else metrics[b]).get(key)) for b in banks]
        vals = [(b, v) for b, v in vals if v is not None]
        vals.sort(key=lambda t: -sign * t[1])
        for i, (b, _) in enumerate(vals):
            ranks[b][name] = i + 1
        for b in banks:
            ranks[b].setdefault(name, None)
    out = {}
    for b in banks:
        pct = [1 - (r - 1) / (len(banks) - 1) for r in ranks[b].values() if r is not None]
        out[b] = {"ranks": ranks[b], "score": 100 * statistics.mean(pct) if pct else None}
    order = sorted(out, key=lambda b: -(out[b]["score"] or -1))
    for i, b in enumerate(order):
        out[b]["overall"] = i + 1
    return out


def compute(data, metrics, adv, vals, prices, ownership) -> dict:
    blk = {b: bank_block(data[b], (m["year"], m["quarter"]), m) for b, m in metrics.items()}
    return {
        "bank": blk,
        "share": market_share(data),
        "stock": stock_stats(prices or {}, vals or {}, ownership or {}),
        "valuation": valuation_extra(vals or {}, metrics),
        "forecast": {b: forecast(data[b], (m["year"], m["quarter"])) for b, m in metrics.items()},
        "score": scorecard(metrics, adv),
        "_vals": vals or {},
    }


# ═════════════════════════ ДҮГНЭЛТ ═════════════════════════════════════════
def highlights(ext: dict, metrics: dict) -> list[str]:
    L = []
    sh = ext["share"].get("banks", {})
    ch = [(b, v["assets_chg"]) for b, v in sh.items() if v.get("assets_chg") is not None]
    if ch:
        up, dn = max(ch, key=lambda t: t[1]), min(ch, key=lambda t: t[1])
        L.append(f"Зах зээлийн хувь (хөрөнгө, {ext['share']['period']}, жилийн өмнөхөөс): хамгийн их өссөн "
                 f"{up[0]} ({up[1] * 100:+.1f} н.х.), хамгийн их буурсан {dn[0]} ({dn[1] * 100:+.1f} н.х.).")
    sp = [(b, v["asset_yield"] - v["funding_cost"]) for b, v in ext["bank"].items()
          if None not in (v["asset_yield"], v["funding_cost"])]
    if sp:
        hi, lo = max(sp, key=lambda t: t[1]), min(sp, key=lambda t: t[1])
        L.append(f"Хүүгийн спред (активын өгөөж − эх үүсвэрийн өртөг): хамгийн өргөн {hi[0]} ({hi[1]:.1%}), "
                 f"хамгийн нарийн {lo[0]} ({lo[1]:.1%}).")
    sc = ext["score"]
    order = sorted(sc, key=lambda b: sc[b]["overall"])
    if order:
        L.append(f"Нэгдсэн скоркард (11 үзүүлэлт): 1-р {order[0]} ({sc[order[0]]['score']:.0f} оноо), "
                 f"2-р {order[1]}, 3-р {order[2]}; сүүлийн байр {order[-1]} ({sc[order[-1]]['score']:.0f}).")
    st = ext["stock"]
    if st:
        r = [(b, v["r_3m"]) for b, v in st.items() if v["r_3m"] is not None]
        if r:
            best, worst = max(r, key=lambda t: t[1]), min(r, key=lambda t: t[1])
            L.append(f"Хувьцаа, 3 сар: хамгийн их өссөн {best[0]} ({best[1]:+.1%}), хамгийн их буурсан "
                     f"{worst[0]} ({worst[1]:+.1%}) — ханшийн өгөгдөл {max(v['last_date'] for v in st.values())} хүртэл.")
    gv = [(b, v["graham_gap"]) for b, v in ext["valuation"].items() if v["graham_gap"] is not None]
    if gv:
        b, x = max(gv, key=lambda t: t[1])
        L.append(f"Graham тоотой харьцуулахад хамгийн их зөрүүтэй: {b} (Graham тоо ханшаас {x:+.0%}).")
    fc = [(b, f) for b, f in ext["forecast"].items() if f and b in ba.MSE_BANKS]
    if fc:
        L.append("Дараагийн улирлын дан ашгийн чиг хандлагын тооцоо (МХБ): "
                 + ", ".join(f"{b} {f0(f['forecast'])} сая ₮" for b, f in fc) + " — жишиг, таамаглал биш.")
    return L


# ═════════════════════════ ГРАФИК ══════════════════════════════════════════
def _signed_hbars(ax, labels, series, fmt="{:+.0%}", clip=None):
    """Эерэг/сөрөг утгатай хэвтээ бүлэглэсэн баганууд; clip-ээс хэтэрсэн утгыг тайраад
    шошгон дээр бодит утгыг бичнэ."""
    import numpy as np
    n = len(series)
    h = .8 / n
    yy = np.arange(len(labels))
    allv = [v for _, vals, _ in series for v in vals if v is not None]
    lim = clip or (max(abs(v) for v in allv) if allv else 1)
    for i, (name, vals, color) in enumerate(series):
        pos = yy - .4 + h * (i + .5)
        shown = [max(-lim, min(lim, v)) if v is not None else 0 for v in vals]
        ax.barh(pos, shown, height=h * .88, color=color, label=name, edgecolor=ar.SURFACE, linewidth=1)
        for p_, v, s in zip(pos, vals, shown):
            txt = fmt.format(v) if v is not None else "—"
            ax.text(s + (lim * .015 if s >= 0 else -lim * .015), p_, txt + ("»" if v is not None and abs(v) > lim else ""),
                    va="center", ha="left" if s >= 0 else "right", fontsize=7, color=ar.INK2)
    ax.axvline(0, color=ar.MUTED, lw=.8)
    ax.set_yticks(yy)
    ax.set_yticklabels(labels)
    ax.invert_yaxis()
    ax.set_xlim(-lim * 1.35, lim * 1.35)
    ax.grid(axis="y", visible=False)


def _trend_grid(data, order, fn, title, path, pct=True):
    plt = ar._plt()
    hist = ar.ratio_history(data, fn, n=10)
    periods = sorted({pp for v in hist.values() for pp, _ in v})
    med = {}
    for pp in periods:
        xs = [x for b in hist for q_, x in hist[b] if q_ == pp and x is not None]
        med[pp] = statistics.median(xs) if xs else None
    fig, axes = plt.subplots(3, 3, figsize=(11, 7.5), sharey=True)
    for ax, b in zip(axes.flat, order):
        ser = hist[b]
        xs_ = [periods.index(pp) for pp, x in ser if x is not None]
        ys_ = [x for _, x in ser if x is not None]
        mx = [i for i, pp in enumerate(periods) if med[pp] is not None]
        ax.plot(mx, [med[periods[i]] for i in mx], color=ar.NEUTRAL, lw=1.5, label="Салбарын медиан")
        ax.plot(xs_, ys_, color=ar.SERIES[0], lw=2, marker="o", ms=3, label="Банк")
        if ys_:
            ax.annotate(f"{ys_[-1]:.1%}" if pct else f"{ys_[-1]:.2f}", (xs_[-1], ys_[-1]), xytext=(4, 4),
                        textcoords="offset points", fontsize=7.5, color=ar.INK2)
        ax.set_title(b, fontsize=9.5)
        ax.set_xticks([0, len(periods) - 1])
        ax.set_xticklabels([periods[0], periods[-1]], fontsize=7.5)
        if pct:
            ar._pctfmt(ax, "y")
    for ax in list(axes.flat)[len(order):]:
        ax.set_visible(False)
    axes.flat[0].legend(loc="upper left", fontsize=7)
    fig.suptitle(title, x=.01, ha="left", fontsize=11, fontweight="bold", color=ar.INK)
    fig.tight_layout(rect=(0, 0, 1, .96))
    return ar._save(fig, path)


def make_charts(data, metrics, adv, ext, out_dir: Path) -> dict[str, list[Path]]:
    plt = ar._plt()
    import numpy as np
    out_dir = Path(out_dir)
    order = sorted(metrics, key=lambda b: -(metrics[b]["assets"] or 0))
    latest = max(m["date"] for m in metrics.values())
    lab = [f"{b} ({metrics[b]['date'][2:]})" if metrics[b]["date"] != latest else b for b in order]
    E = ext["bank"]
    C: dict[str, list[Path]] = {}

    # A. өсөлт — баланс ба орлого тусдаа самбар
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 5.4), gridspec_kw={"wspace": .55})
    _signed_hbars(a1, lab, [("Зээл", [E[b]["g_loans"] for b in order], ar.SERIES[0]),
                            ("Харилцах+хадгаламж", [E[b]["g_deposits"] for b in order], ar.SERIES[1]),
                            ("Өөрийн хөрөнгө", [E[b]["g_equity"] for b in order], ar.SERIES[2])], clip=.8)
    a1.set_title("Балансын жилийн өсөлт")
    ar._pctfmt(a1)
    a1.legend(loc="upper center", bbox_to_anchor=(.42, -.07), ncol=3, fontsize=8)
    _signed_hbars(a2, lab, [("Хүүгийн цэвэр орлого", [E[b]["g_nii"] for b in order], ar.SERIES[0]),
                            ("Хүүгийн бус зардал", [E[b]["g_opex"] for b in order], ar.SERIES[1]),
                            ("Цэвэр ашиг", [E[b]["g_profit"] for b in order], ar.SERIES[2])], clip=.8)
    a2.set_title("Орлого, зардал, ашгийн өсөлт (YTD, YoY)")
    ar._pctfmt(a2)
    a2.legend(loc="upper center", bbox_to_anchor=(.42, -.07), ncol=3, fontsize=8)
    fig.subplots_adjust(left=.14, right=.96, top=.9, bottom=.16)
    C["growth"] = [ar._save(fig, out_dir / "ext_growth.png", "» = 80%-иас давсан утга (тайрсан).")]

    # A. зах зээлийн хувь
    sh = ext["share"]
    if sh:
        bs = sorted(sh["banks"], key=lambda b: -(sh["banks"][b]["assets"] or 0))
        fig, ax = plt.subplots(figsize=(9, 4.8))
        ar._hbars(ax, bs, [("Хөрөнгө", [sh["banks"][b]["assets"] for b in bs], ar.SERIES[0]),
                           ("Зээл", [sh["banks"][b]["loans"] for b in bs], ar.SERIES[1]),
                           ("Харилцах+хадгаламж", [sh["banks"][b]["deposits"] for b in bs], ar.SERIES[2])])
        for i, b in enumerate(bs):
            c = sh["banks"][b]["assets_chg"]
            if c is not None:
                ax.text(ax.get_xlim()[1], i, f"{c * 100:+.1f} н.х.", ha="right", va="center", fontsize=7.5,
                        color=ar.SERIES[0] if c >= 0 else ar.CRITICAL)
        ax.set_title(f"Зах зээлийн хувь, {sh['period']} (баруун: хөрөнгийн хувийн жилийн өөрчлөлт)")
        ar._pctfmt(ax)
        ax.legend(loc="upper center", bbox_to_anchor=(.5, -.08), ncol=3)
        fig.tight_layout()
        C["share"] = [ar._save(fig, out_dir / "ext_market_share.png")]

    # B. активын бүтэц
    fig, ax = plt.subplots(figsize=(9, 4.8))
    parts = [("Цэвэр зээл", "mix_loans", ar.SERIES[0]), ("Хөрөнгө оруулалт", "mix_investments", ar.SERIES[1]),
             ("Банк, СБ-д байршуулсан", "mix_placements", ar.SERIES[2]), ("Мөнгө", "mix_cash", ar.SERIES[3]),
             ("Бусад", "mix_other", ar.NEUTRAL)]
    left = np.zeros(len(order))
    for name, key, color in parts:
        v = np.array([max(E[b][key] or 0, 0) for b in order])
        ax.barh(lab, v, left=left, color=color, label=name, edgecolor=ar.SURFACE, linewidth=2, height=.62)
        for i, (l_, vv) in enumerate(zip(left, v)):
            if vv >= .07:
                ax.text(l_ + vv / 2, i, f"{vv:.0%}", ha="center", va="center", fontsize=7.5,
                        color="#ffffff" if color in (ar.SERIES[0], ar.SERIES[1]) else ar.INK)
        left += v
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    ar._pctfmt(ax)
    ax.grid(axis="y", visible=False)
    ax.set_title("Активын бүтэц (нийт хөрөнгөд эзлэх хувь)")
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.08), ncol=5)
    fig.tight_layout()
    C["structure"] = [ar._save(fig, out_dir / "ext_asset_mix.png", "Голомт, Хаан банкны 'Бусад'-д хуудсанд тусдаа мөргүй зүйлс орно.")]

    # C. өгөөж, өртөг, спред
    fig, ax = plt.subplots(figsize=(9, 5))
    ar._hbars(ax, lab, [("Хүү өгөх активын өгөөж", [E[b]["asset_yield"] for b in order], ar.SERIES[0]),
                        ("Эх үүсвэрийн өртөг", [E[b]["funding_cost"] for b in order], ar.SERIES[1]),
                        ("NIM (хүү өгөх активт)", [E[b]["nim_iea"] for b in order], ar.SERIES[2])])
    ax.set_title("Өгөөж, эх үүсвэрийн өртөг, маржин (жилийн хэмжээст, дундаж үлдэгдэлд)")
    ar._pctfmt(ax)
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.08), ncol=3)
    fig.tight_layout(rect=(0, .03, 1, 1))
    C["yield"] = [ar._save(fig, out_dir / "ext_yield_cost.png",
                           "Хүү өгөх актив = банкинд байршуулсан + хөрөнгө оруулалт + цэвэр зээл; эх үүсвэр = харилцах, хадгаламж + банк/СБ + бусад эх үүсвэр.")]

    # D. ашгийн чанар
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 5), gridspec_kw={"wspace": .55})
    ar._hbars(a1, lab, [("Хүүгийн бус орлогын хувь", [E[b]["nonint_share"] for b in order], ar.SERIES[0])])
    a1.set_title("Хүүгийн бус орлогын хувь")
    ar._pctfmt(a1)
    ar._hbars(a2, lab, [("Эрсдэлийн сан / PPOP", [E[b]["prov_ppop"] for b in order], ar.SERIES[1])], ref=.5, ref_label="50%")
    a2.set_title("Эрсдэлийн сан / сангийн өмнөх ашиг")
    ar._pctfmt(a2)
    fig.subplots_adjust(left=.14, right=.98, top=.9, bottom=.08)
    C["quality"] = [ar._save(fig, out_dir / "ext_earnings_quality.png")]

    # F. скоркард heatmap
    sc = ext["score"]
    so = sorted(sc, key=lambda b: sc[b]["overall"])
    names = [n for n, *_ in SCORE_METRICS]
    M = np.array([[sc[b]["ranks"][n] if sc[b]["ranks"][n] is not None else np.nan for n in names] for b in so])
    from matplotlib.colors import LinearSegmentedColormap
    cmap = LinearSegmentedColormap.from_list("blue", ["#184f95", "#3987e5", "#86b6ef", "#cde2fb"])
    fig, ax = plt.subplots(figsize=(11.5, 5.2))
    ax.imshow(M, cmap=cmap, vmin=1, vmax=len(so), aspect="auto")
    for i in range(len(so)):
        for j in range(len(names)):
            v = M[i, j]
            ax.text(j, i, "—" if np.isnan(v) else f"{int(v)}", ha="center", va="center", fontsize=8,
                    color="#ffffff" if (not np.isnan(v) and v <= len(so) / 2) else ar.INK)
    ax.set_xticks(range(len(names)))
    ax.set_xticklabels(names, rotation=30, ha="right", fontsize=8)
    ax.set_yticks(range(len(so)))
    ax.set_yticklabels([f"{sc[b]['overall']}. {b}  ({sc[b]['score']:.0f})" for b in so], fontsize=8.5)
    ax.grid(False)
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_title("Нэгдсэн скоркард: үзүүлэлт бүрийн эрэмбэ (1 = хамгийн сайн; хаалтанд нийлмэл оноо 0–100)")
    fig.tight_layout()
    C["score"] = [ar._save(fig, out_dir / "ext_scorecard.png")]

    # G. хувьцаа
    st = ext["stock"]
    if st:
        bs = sorted(st, key=lambda b: -(st[b]["r_1y"] or -9))
        fig, ax = plt.subplots(figsize=(9, 4.4))
        _signed_hbars(ax, bs, [("1 сар", [st[b]["r_1m"] for b in bs], ar.SERIES[0]),
                               ("3 сар", [st[b]["r_3m"] for b in bs], ar.SERIES[1]),
                               ("6 сар", [st[b]["r_6m"] for b in bs], ar.SERIES[2]),
                               ("1 жил", [st[b]["r_1y"] for b in bs], ar.SERIES[3])])
        ax.set_title(f"Хувьцааны өгөөж (ханш {max(v['last_date'] for v in st.values())} хүртэл)")
        ar._pctfmt(ax)
        ax.legend(loc="upper center", bbox_to_anchor=(.5, -.08), ncol=4)
        fig.tight_layout()
        p1 = ar._save(fig, out_dir / "ext_stock_returns.png")
        fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 4.6), gridspec_kw={"wspace": .35})
        pts = [(b, st[b]["vol_1y"], st[b]["r_1y"]) for b in bs if None not in (st[b]["vol_1y"], st[b]["r_1y"])]
        for b, xv, yv in pts:
            a1.scatter(xv, yv, s=60, color=ar.SERIES[0], edgecolor=ar.SURFACE, linewidth=2, zorder=3)
        a1.axhline(0, color=ar.MUTED, lw=.8)
        ar._pctfmt(a1, "x")
        ar._pctfmt(a1, "y")
        a1.set_xlabel("Жилийн хэлбэлзэл (өдрийн өгөөжийн стандарт хазайлт)")
        a1.set_ylabel("1 жилийн өгөөж")
        a1.set_title("Эрсдэл ба өгөөж")
        a1.margins(.2)
        ar._hbars(a2, bs, [("Өдрийн дундаж арилжаа, сая ₮", [st[b]["adtv_3m"] for b in bs], ar.SERIES[1])], fmt="{:,.1f}")
        a2.set_title("Хөрвөх чадвар: өдрийн дундаж арилжаа (3 сар)")
        fig.subplots_adjust(left=.07, right=.98, top=.88, bottom=.15)
        ar._label_points(a1, pts)
        p2 = ar._save(fig, out_dir / "ext_stock_risk_liquidity.png")
        C["stock"] = [p1, p2]

    # H. үнэлгээ: P/E ба ашгийн өсөлт
    va = ext["valuation"]
    if va:
        fig, ax = plt.subplots(figsize=(7.5, 4.6))
        pts = []
        from_vals = ext.get("_vals", {})
        for b in va:
            pe, g = (from_vals.get(b) or {}).get("pe"), metrics[b]["profit_ytd_yoy"]
            if pe and g is not None:
                pts.append((b, g, pe))
                ax.scatter(g, pe, s=60, color=ar.SERIES[0], edgecolor=ar.SURFACE, linewidth=2, zorder=3)
        ar._pctfmt(ax, "x")
        ax.set_xlabel("Цэвэр ашгийн жилийн өсөлт (YTD)")
        ax.set_ylabel("P/E (TTM)")
        ax.set_title("Үнэлгээ ба өсөлт: баруун доод = өсөлттэй атлаа хямд")
        ax.margins(.2)
        fig.tight_layout()
        ar._label_points(ax, pts)
        C["valuation"] = [ar._save(fig, out_dir / "ext_pe_growth.png")]

    # I. прогноз — small multiples
    fc = ext["forecast"]
    fig, axes = plt.subplots(3, 3, figsize=(11, 7.5))
    for ax, b in zip(axes.flat, order):
        f = fc.get(b)
        if not f:
            ax.set_visible(False)
            continue
        hx = [h for h in f["history"] if h[1] is not None]
        xs = list(range(len(hx)))
        ax.bar(xs, [v / 1000 for _, v in hx], color=ar.SERIES[0], width=.62, edgecolor=ar.SURFACE)
        n = len(hx)
        ax.bar([n], [f["forecast"] / 1000], color=ar.SURFACE, edgecolor=ar.SERIES[1], hatch="///", width=.62, lw=1.2)
        ax.errorbar([n], [f["forecast"] / 1000], yerr=[[(f["forecast"] - f["low"]) / 1000], [(f["high"] - f["forecast"]) / 1000]],
                    color=ar.SERIES[1], capsize=3, lw=1)
        ax.set_xticks([0, n])
        ax.set_xticklabels([hx[0][0], f["next"] + "*"], fontsize=7.5)
        ax.set_title(b, fontsize=9.5)
        ax.axhline(0, color=ar.MUTED, lw=.6)
        ax.text(n, f["high"] / 1000, f" {f['forecast'] / 1000:,.1f}", fontsize=7, color=ar.INK2, va="bottom")
    for ax in list(axes.flat)[len(order):]:
        ax.set_visible(False)
    fig.suptitle("Улирлын дан цэвэр ашиг (тэрбум ₮) ба дараагийн улирлын чиг хандлагын тооцоо (*, ±1 хазайлт)",
                 x=.01, ha="left", fontsize=11, fontweight="bold", color=ar.INK)
    fig.tight_layout(rect=(0, .03, 1, .96))
    C["forecast"] = [ar._save(fig, out_dir / "ext_forecast.png",
                              "Сүүлийн 6 улирлын шугаман чиг хандлага — жишиг тооцоо, хөрөнгө оруулалтын зөвлөмж биш.")]

    # J. чиг хандлагууд
    def mk(key):
        def fn(bd, k):
            if not bd.bal.get(k) or not bd.inc.get(k):
                return None
            try:
                return ba.compute_bank(bd, k)[key]
            except Exception:  # noqa: BLE001
                return None
        return fn
    C["trends"] = [
        _trend_grid(data, order, mk("nim"), "NIM (жилийн хэмжээст) — банк бүр ба салбарын медиан", out_dir / "ext_trend_nim.png"),
        _trend_grid(data, order, mk("cost_income"), "Зардал / орлого — банк бүр ба салбарын медиан", out_dir / "ext_trend_cost_income.png"),
        _trend_grid(data, order, mk("equity_assets"), "Өөрийн хөрөнгө / нийт хөрөнгө — банк бүр ба салбарын медиан", out_dir / "ext_trend_equity_assets.png"),
        _trend_grid(data, order, mk("ltd"), "Зээл / (харилцах + хадгаламж) — банк бүр ба салбарын медиан", out_dir / "ext_trend_ltd.png"),
    ]
    return C


# ═════════════════════════ WORD БҮЛГҮҮД ═══════════════════════════════════
def sections(ext: dict, metrics: dict, charts: dict, vals: dict) -> list[dict]:
    order = sorted(metrics, key=lambda b: -(metrics[b]["assets"] or 0))
    E = ext["bank"]

    def row(b, keys, f=fp, sign=False):
        return [b] + [f(E[b][k], sign) if f is fp else f(E[b][k]) for k in keys]

    S = []
    sh = ext["share"]
    share_rows = []
    if sh:
        for b in sorted(sh["banks"], key=lambda b: -(sh["banks"][b]["assets"] or 0)):
            v = sh["banks"][b]

            def pp(x):
                return "—" if x is None else f"{x * 100:+.1f}"
            share_rows.append([b, fp(v["assets"]), pp(v["assets_chg"]), fp(v["loans"]), pp(v["loans_chg"]),
                               fp(v["deposits"]), pp(v["deposits_chg"])])
    S.append({"title": "Өсөлт ба зах зээлийн хувь",
              "intro": "Банк тус бүрийн сүүлийн тайлангаар; YTD үзүүлэлтийг өмнөх оны мөн үетэй харьцуулав.",
              "tables": [("Жилийн (YoY) ба улирлын (QoQ) өсөлт",
                          ["Банк", "Хөрөнгө", "Зээл", "Хадг.", "ӨХ", "ЦХО", "Зардал", "Ашиг", "Хөр. QoQ", "Зээл QoQ", "Хөр. 2ж CAGR"],
                          [row(b, ["g_assets", "g_loans", "g_deposits", "g_equity", "g_nii", "g_opex", "g_profit",
                                   "q_assets", "q_loans", "cagr2_assets"], sign=True) for b in order])]
              + ([(f"Зах зээлийн хувь, {sh['period']} (н.х. = жилийн өмнөхөөс нэгж хувийн өөрчлөлт)",
                   ["Банк", "Хөрөнгө", "н.х.", "Зээл", "н.х.", "Хадг.", "н.х."], share_rows)] if sh else []),
              "charts": charts.get("growth", []) + charts.get("share", [])})
    S.append({"title": "Балансын ба орлогын бүтэц",
              "tables": [("Активын бүтэц ба орлогын төрөлжилт",
                          ["Банк", "Цэвэр зээл", "Хөрөнгө оруулалт", "Байршуулсан", "Мөнгө", "Бусад", "Хадг./хөр.", "Хүүгийн бус орл. хувь"],
                          [row(b, ["mix_loans", "mix_investments", "mix_placements", "mix_cash", "mix_other",
                                   "deposits_assets", "nonint_share"]) for b in order])],
              "charts": charts.get("structure", [])})
    S.append({"title": "Өгөөж, эх үүсвэрийн өртөг ба маржин",
              "tables": [("Дундаж үлдэгдэлд, жилийн хэмжээст",
                          ["Банк", "Активын өгөөж", "Эх үүсвэрийн өртөг", "Спред", "NIM (хүү өгөх активт)", "NIM (нийт хөрөнгөд)"],
                          [[b, fp(E[b]["asset_yield"]), fp(E[b]["funding_cost"]),
                            fp(E[b]["asset_yield"] - E[b]["funding_cost"]) if None not in (E[b]["asset_yield"], E[b]["funding_cost"]) else "—",
                            fp(E[b]["nim_iea"]), fp(metrics[b]["nim"])] for b in order])],
              "charts": charts.get("yield", [])})
    S.append({"title": "Ашгийн чанар ба үр ашиг",
              "tables": [("Ашгийн чанар", ["Банк", "Сан/ЦХО", "Сан/PPOP", "Татварын бодит %", "Зардал/хөр.", "Орлого/хөр.",
                                           "ТӨ ROA", "TTM ROE", "Улирлын ROE", "Ашгийн CV"],
                          [[b, fp(E[b]["prov_nii"]), fp(E[b]["prov_ppop"]), fp(E[b]["eff_tax"]), fp(E[b]["opex_assets"]),
                            fp(E[b]["rev_assets"]), fp(E[b]["roa_pretax"]), fp(E[b]["roe_ttm"]), fp(E[b]["roe_quarter"]),
                            "—" if E[b]["profit_cv"] is None else f"{E[b]['profit_cv']:.2f}"] for b in order])],
              "notes": ["Ашгийн CV = сүүлийн 8 улирлын дан ашгийн стандарт хазайлт / дундаж (бага = тогтвортой).",
                        "Улирлын ROE = зөвхөн тайлант улирлын ашиг × 4 / өөрийн хөрөнгө — хамгийн сүүлийн хурдыг харуулна."],
              "charts": charts.get("quality", [])})
    S.append({"title": "Өргөтгөсөн эрсдэлийн үзүүлэлтүүд",
              "tables": [("Активын чанар, хөрвөх чадвар, капитал",
                          ["Банк", "Нөөцийн харьцаа", "NPL/ӨХ", "Цэвэр NPL/ӨХ", "Мөнгө/хадг.", "Зээл/эх үүсвэр",
                           "Бөөний эх үүсвэр", "ӨХ/зээл"],
                          [row(b, ["reserve_ratio", "npl_equity", "net_npl_equity", "cash_deposits", "loans_funding",
                                   "wholesale_share", "equity_loans"]) for b in order])],
              "notes": ["Нөөцийн харьцаа = зээлийн эрсдэлийн сан / нийт зээл. Цэвэр NPL/ӨХ = (NPL − эрсдэлийн сан) / ӨХ; "
                        "сөрөг бол сан чанаргүй зээлээс их.",
                        "Бөөний эх үүсвэр = (банк, СБ-аас татсан + бусад эх үүсвэр) / нийт өр төлбөр."]})
    sc = ext["score"]
    so = sorted(sc, key=lambda b: sc[b]["overall"])
    names = [n for n, *_ in SCORE_METRICS]
    S.append({"title": "Нэгдсэн скоркард",
              "intro": "11 үзүүлэлтээр банк бүрийг эрэмбэлж (1 = хамгийн сайн), эрэмбийг 0–100 оноо болгон дундажлав. "
                       "Мэдээлэлгүй үзүүлэлтийг тооцоонд оруулаагүй. Хуучин огнооны тайлантай банкуудыг тухайн огноогоор нь тооцов.",
              "tables": [("Эрэмбэ", ["#", "Банк", "Оноо"] + names,
                          [[str(sc[b]["overall"]), b, f"{sc[b]['score']:.0f}"] + [str(sc[b]["ranks"][n] or "—") for n in names]
                           for b in so])],
              "charts": charts.get("score", [])})
    st = ext["stock"]
    if st:
        bs = sorted(st, key=lambda b: -(st[b]["r_1y"] or -9))

        def f1(x):
            return "—" if x is None else f"{x:,.1f}"
        S.append({"title": "Хувьцааны гүйцэтгэл, эрсдэл ба хөрвөх чадвар (МХБ-ийн 5 банк)",
                  "intro": f"Ханшийн өгөгдөл {max(v['last_date'] for v in st.values())} хүртэл.",
                  "tables": [("Өгөөж", ["Банк", "Ханш", "1 сар", "3 сар", "6 сар", "1 жил", "Оны эхнээс"],
                              [[b, f0(st[b]["price"]), fp(st[b]["r_1m"], True), fp(st[b]["r_3m"], True), fp(st[b]["r_6m"], True),
                                fp(st[b]["r_1y"], True), fp(st[b]["r_ytd"], True)] for b in bs]),
                             ("Эрсдэл ба хөрвөх чадвар", ["Банк", "Хэлбэлзэл (жил)", "Их уналт (1ж)", "52 д.х. оргилоос",
                                                          "Өдрийн дундаж арилжаа, сая ₮", "Арилжаатай өдөр (ажлын өдөрт)", "Эргэлт (1ж)"],
                              [[b, fp(st[b]["vol_1y"]), fp(st[b]["max_dd_1y"]), fp(st[b]["from_high"], True), f1(st[b]["adtv_3m"]),
                                fp(st[b]["days_traded_3m"]), fp(st[b]["turnover_1y"])] for b in bs])],
                  "notes": ["Эргэлт = жилийн нийт арилжсан хувьцаа / чөлөөт эргэлтийн хувьцаа (ҮЦТХТ байвал, эс бөгөөс нийт хувьцаа)."],
                  "charts": charts.get("stock", [])})
    va = ext["valuation"]
    if va:
        S.append({"title": "Өргөтгөсөн үнэлгээ (МХБ-ийн 5 банк)",
                  "tables": [("Үнэлгээний харьцаа", ["Банк", "E/P", "ROE / P/B", "Зах зээлийн үнэлгээ / хөрөнгө",
                                                     "Graham тоо, ₮", "Graham ханшаас", "PEG"],
                              [[b, fp(va[b]["earnings_yield"]), "—" if va[b]["roe_pb"] is None else f"{va[b]['roe_pb']:.1%}",
                                fp(va[b]["mcap_assets"]), f0(va[b]["graham"]), fp(va[b]["graham_gap"], True),
                                "—" if va[b]["peg"] is None else f"{va[b]['peg']:.2f}"] for b in va])],
                  "notes": ["Graham тоо = √(22.5 × EPS × BVPS). ROE / P/B — өндөр бол өгөөжтэй харьцуулахад хямд. "
                            "PEG = P/E / (ашгийн өсөлт × 100); зөвхөн эерэг өсөлттэй үед."],
                  "charts": charts.get("valuation", [])})
    fc = ext["forecast"]
    S.append({"title": "Прогноз: дараагийн улирлын ашгийн чиг хандлага",
              "intro": "Сүүлийн 6 улирлын дан ашигт шугаман чиг хандлага тааруулж дараагийн улирлыг тооцов; муж нь "
                       "үлдэгдлийн ±1 стандарт хазайлт. Энэ нь жишиг тооцоо бөгөөд нэг удаагийн орлого, бодлогын өөрчлөлтийг тусгахгүй.",
              "tables": [("Тооцоо (сая ₮)", ["Банк", "Улирал", "Тооцоо", "Доод", "Дээд", "Улирлын чиг", "Жилийн ашиг (тооцоо)"],
                          [[b, fc[b]["next"], f0(fc[b]["forecast"]), f0(fc[b]["low"]), f0(fc[b]["high"]),
                            f"{fc[b]['trend_per_q']:+,.0f}", f0(fc[b]["fy_estimate"])] for b in order if fc.get(b)])],
              "charts": charts.get("forecast", [])})
    S.append({"title": "Үзүүлэлтүүдийн чиг хандлага (10 улирал)", "charts": charts.get("trends", [])})
    return S


def _end_labels(ax, ends, fmt, log=False, min_gap_pt=11):
    """Шугамын төгсгөлийн утгын шошгыг босоо чиглэлд давхцуулахгүй байрлуулна."""
    fig = ax.figure
    fig.canvas.draw()
    k = fig.dpi / 72.0
    pts = sorted(((b, x, y, ax.transData.transform((ax.xaxis.convert_units(x), y))[1]) for b, x, y in ends),
                 key=lambda t: t[3])
    placed = []
    for b, x, y, py in pts:
        ty = py if not placed else max(py, placed[-1] + min_gap_pt * k)
        placed.append(ty)
        ax.annotate(fmt.format(y), (x, y), xytext=(6, (ty - py) / k), textcoords="offset points",
                    fontsize=7.5, color=ar.INK2, va="center")


# ═════════════════════════ ЭЗЭМШЛИЙН ГРАФИК (5%-иас дээш / доош тусад нь) ════════
def ownership_charts(vals: dict, ownership: dict, out_dir: Path) -> list[Path]:
    plt = ar._plt()
    import numpy as np
    from matplotlib.dates import DateFormatter
    out_dir = Path(out_dir)
    bs = [b for b in sorted(ownership, key=lambda b: -(ownership[b].get("dom_pct") or 0)) if ownership[b].get("over5")]
    if not bs:
        return []
    paths = []

    # 1. Бүтэц: 5%-иас ДЭЭШ ба ДООШ (+ төрийн) — 100% давхарласан
    fig, ax = plt.subplots(figsize=(10, 3.8))
    y = np.arange(len(bs))
    left = np.zeros(len(bs))
    groups = [("Төрийн мэдэлд", lambda o: o.get("state_pct") or 0, None, ar.SERIES[2]),
              ("5%-иас ДЭЭШ эзэмшигчид", lambda o: o["over5"]["pct"] or 0, lambda o: o["over5"]["holders"], ar.SERIES[0]),
              ("5%-иас ДООШ эзэмшигчид", lambda o: o["under5"]["pct"] or 0, lambda o: o["under5"]["holders"], ar.SERIES[1])]
    for name, val, cnt, color in groups:
        v = np.array([val(ownership[b]) for b in bs])
        if not v.any():
            continue
        ax.barh(y, v, left=left, color=color, label=name, edgecolor=ar.SURFACE, linewidth=2, height=.6)
        for i, b in enumerate(bs):
            if v[i] <= 0:
                continue
            txt = f"{v[i]:.1%}" + (f" · {cnt(ownership[b]):,.0f} эзэмшигч" if cnt else "")
            if v[i] >= .25:
                ax.text(left[i] + v[i] / 2, i, txt, ha="center", va="center", fontsize=8, color="#ffffff")
            else:   # нарийн хэсэг — баруун талд гадна
                ax.text(1.01, i, txt, ha="left", va="center", fontsize=8, color=ar.INK2, transform=ax.get_yaxis_transform())
        left += v
    ax.set_yticks(y)
    ax.set_yticklabels(bs)
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    ar._pctfmt(ax)
    ax.grid(axis="y", visible=False)
    as_of = max(o["as_of"] for o in ownership.values())
    ax.set_title(f"Хувьцаа эзэмшлийн бүтэц: 5%-иас дээш ба доош эзэмшигчид ({as_of})")
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.12), ncol=3)
    fig.subplots_adjust(left=.12, right=.78, top=.88, bottom=.24)
    paths.append(ar._save(fig, out_dir / "own_structure.png", "Эх сурвалж: ҮЦТХТ-ийн үнэт цаасны төвлөрлийн судалгаа."))

    # 2. Түүх: (а) 5%-иас ДООШ эзэмшигчдийн тоо, (б) 5%-иас ДЭЭШ эзэмшигчдийн эзлэх хувь
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 4.6), gridspec_kw={"wspace": .3})
    ends1, ends2 = [], []
    for i, b in enumerate(bs):
        h = [x for x in ownership[b]["history"] if x["under5_holders"]]
        if not h:
            continue
        d = [date.fromisoformat(x["date"]) for x in h]
        col = ar.SERIES[i % len(ar.SERIES)]
        a1.plot(d, [x["under5_holders"] for x in h], color=col, lw=2, marker="o", ms=3, label=b)
        a2.step(d, [x["over5_pct"] for x in h], where="post", color=col, lw=2, label=b)
        ends1.append((b, d[-1], h[-1]["under5_holders"]))
        ends2.append((b, d[-1], h[-1]["over5_pct"]))
    a1.set_yscale("log")
    from matplotlib.ticker import FixedLocator, FuncFormatter, NullFormatter
    a1.yaxis.set_major_locator(FixedLocator([1000, 2000, 5000, 10000, 20000, 50000, 100000]))
    a1.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
    a1.yaxis.set_minor_formatter(NullFormatter())
    _end_labels(a1, ends1, "{:,.0f}", log=True)
    _end_labels(a2, ends2, "{:.1%}")
    a1.set_title("5%-иас ДООШ эзэмшигчдийн тоо (лог тэнхлэг)")
    a2.set_title("5%-иас ДЭЭШ эзэмшигчдийн эзлэх хувь")
    ar._pctfmt(a2, "y")
    for ax in (a1, a2):
        ax.xaxis.set_major_formatter(DateFormatter("%Y/%m"))
        ax.tick_params(axis="x", rotation=30)
        ax.margins(x=.08)
    a1.legend(loc="upper center", bbox_to_anchor=(1.1, -.2), ncol=len(bs))
    fig.subplots_adjust(left=.07, right=.95, top=.9, bottom=.25)
    paths.append(ar._save(fig, out_dir / "own_history.png"))
    return paths
