#!/usr/bin/env python3
"""
advanced_ratios.py — банкны НЭМЭЛТ харьцаа үзүүлэлт ба тэдгээрийн графикууд.

bank_analysis.BankData (workbook-ийн түүхий баланс/орлого) дээр ажиллана. Тооцох зүйлс:

 Зээлийн эрсдэл   Хугацаа хэтэрсэн + NPL харьцаа, Texas харьцаа, хоцрогдсон NPL
                  (NPL_t / жилийн өмнөх нийт зээл), NPL-ийн улирлын өөрчлөлт, зээлийн
                  эрсдэлийн өртөг (CoR), стресс тест (NPL ×2, ×3)
 Эх үүсвэр,       Харилцах+хадгаламжийн эзлэх хувь, харилцахын хувь (CASA), банк/СБ-аас
 хөрвөх чадвар    татсан эх үүсвэр, хөрвөх хөрөнгө / нийт хөрөнгө, хөрвөх хөрөнгө / хадгаламж
 Ашгийн чанар     DuPont (ROA-г дундаж хөрөнгөд: ЦХО + хүүгийн бус орлого − зардал −
                  эрсдэлийн сан − татвар ± бусад), хөшүүрэг, дундаж ӨХ-ээр ROE, сангийн
                  өмнөх ашиг (PPOP), Jaws
 Капитал          Дотоод капитал үүсгэлт (ICG = ROE, ногдол ашиггүй гэж) ба хөрөнгийн
                  өсөлтийн зөрүү, биет өөрийн хөрөнгө / хөрөнгө
 Үнэлгээ          Үндэслэлтэй P/B = (ROE − g) / (COE − g)  (COE, g — таамаглал, тохируулна)
 Салбар           HHI (хөрөнгө, зээл, хадгаламж) улирал бүрээр

Мэдээлэл байхгүй мөрийг (жиш. Голомт, Хаан банкны хуудсанд "хугацаа хэтэрсэн зээл"
мөр алга) ТЭГ гэж үзэхгүй — харьцаа None ("—") болно.

Графикийн дүрэм (dataviz): банкийг ӨНГӨӨР биш тэнхлэгийн нэрээр ялгана (9 банк >
8 өнгө); өнгө зөвхөн бүрэлдэхүүн хэсэг/цувралд, тогтмол дарааллаар; нэг тэнхлэг;
нимгэн зураас, бүдэг тор; утгын шошго (3 өнгө дэвсгэртэй 3:1-ээс бага ялгаатай тул).
"""
from __future__ import annotations

import statistics
from pathlib import Path

import bank_analysis as ba

# ── Тохиргоо (CLI-аас өөрчилнө) ───────────────────────────────────────────
DEFAULTS = {
    "coe": 0.20,       # хувьцааны өгөөжийн шаардлага — ТААМАГЛАЛ
    "growth": 0.06,    # урт хугацааны өсөлт — ТААМАГЛАЛ
    "lgd": 0.50,       # стресс тестийн алдагдлын хувь (шинэ NPL-ийн)
    "texas_alert": 0.40,
    "gap_alert": 0.15,  # хөрөнгийн өсөлт − ICG
}

EXTRA_BAL = {
    "placements": ["банк, сб-д байршуулсан хөрөнгө", "банк байгууллагад байршуулсан",
                   "банк, санхүүгийн байгууллагад байршуулсан хөрөнгө"],
    "investments": ["хөрөнгө оруулалт", "хөрөнгө оруулалт (санхүүгийн)"],
    "bank_borrow": ["банк, сб-аас татсан эх үүсвэр", "банкнаас татсан эх үүсвэр"],
    "other_src": ["бусад эх үүсвэр"],
    "intangible": ["биет бус хөрөнгө"],
}

# ── Өнгө (dataviz reference palette, light) ───────────────────────────────
SURFACE = "#fcfcfb"
INK, INK2, MUTED, GRID = "#0b0b0b", "#52514e", "#898781", "#e8e7e2"
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
NEUTRAL = "#b9b8b2"
CRITICAL = "#d03b3b"


# ── Тооцоо ────────────────────────────────────────────────────────────────
def _extra(bd: ba.BankData, k, key):
    """BankData-д ороогүй нэмэлт баланс мөр (нэрээр)."""
    ws_vals = getattr(bd, "_extra_cache", None)
    if ws_vals is None:
        bd._extra_cache = ws_vals = {}
    return ws_vals.get(k, {}).get(key)


def attach_extra(wb, data: dict[str, ba.BankData]):
    """Workbook-оос EXTRA_BAL мөрүүдийг уншиж BankData-д хавсаргана."""
    for b, bd in data.items():
        ws = wb[f"{b} баланс"]
        bd._extra_cache = bd._read(ws, EXTRA_BAL)


def _sd(a, b):
    return None if a is None or b in (None, 0) else a / b


def _avg(a, b):
    return a if b is None else (b if a is None else (a + b) / 2)


def compute_advanced(bd: ba.BankData, k, m: dict, p: dict) -> dict:
    y, q = k
    ann = ba.ANNUAL[q]
    A, E, L = m["assets"], m["equity"], m["liabilities"]
    gross, npl = bd.gross_loans_npl(k)
    prov = abs(bd.b(k, "provision") or 0) if bd.b(k, "provision") is not None else None
    overdue = bd.b(k, "loans_watch")
    ystart = (y - 1, 4)
    avgA, avgE = _avg(A, bd.b(ystart, "assets")), _avg(E, bd.b(ystart, "equity"))

    nii = bd.ytd(k, "nii")
    ninc, nexp = bd.ytd(k, "nonint_income"), bd.ytd(k, "nonint_expense")
    pexp, pbt, np_ = bd.ytd(k, "provision_exp"), bd.ytd(k, "pbt"), bd.ytd(k, "net_profit")
    tax = pbt - np_ if None not in (pbt, np_) else None

    def a_(x):
        return x * ann / avgA if (x is not None and avgA) else None
    dup = {"nii": a_(nii), "nonint_income": a_(ninc), "opex": a_(nexp), "provisions": a_(pexp),
           "tax": a_(tax), "roa": a_(np_)}
    if None not in dup.values():
        dup["other"] = dup["roa"] - (dup["nii"] + dup["nonint_income"] - dup["opex"] - dup["provisions"] - dup["tax"])
    ppop = (nii + ninc - nexp) * ann if None not in (nii, ninc, nexp) else None

    dep = m["deposits"]
    cur, donly = bd.b(k, "current_acc"), bd.b(k, "deposits_only")
    bank_b, oth = _extra(bd, k, "bank_borrow"), _extra(bd, k, "other_src")
    cash, plc, inv = bd.b(k, "cash"), _extra(bd, k, "placements"), _extra(bd, k, "investments")
    liquid = sum(x for x in (cash, plc, inv) if x is not None) if cash is not None else None
    intang = _extra(bd, k, "intangible")

    npl_prev_q = bd.gross_loans_npl(ba.prev_q(y, q))[1]
    gross_ly = bd.gross_loans_npl((y - 1, q))[0]
    lgd = p["lgd"]

    def stress(mult):
        if None in (npl, E, A):
            return None
        loss = (mult - 1) * npl * lgd
        return (E - loss) / (A - loss)

    # Jaws: YTD орлого ба зардлын жилийн өсөлтийн зөрүү
    rev, rev_ly = (None if None in (nii, ninc) else nii + ninc), None
    nii_ly, ninc_ly, nexp_ly = bd.ytd((y - 1, q), "nii"), bd.ytd((y - 1, q), "nonint_income"), bd.ytd((y - 1, q), "nonint_expense")
    if None not in (nii_ly, ninc_ly):
        rev_ly = nii_ly + ninc_ly
    jaws = (ba.growth(rev, rev_ly) - ba.growth(nexp, nexp_ly)) if None not in (
        ba.growth(rev, rev_ly), ba.growth(nexp, nexp_ly)) else None

    return {
        "bank": bd.bank, "date": m["date"],
        # зээлийн эрсдэл
        "npl_ratio": m["npl_ratio"],
        "overdue_npl_ratio": _sd(npl + overdue, gross) if None not in (npl, overdue) else None,
        "texas": _sd(npl, (E or 0) + (prov or 0)) if npl is not None and E is not None else None,
        "npl_lagged": _sd(npl, gross_ly),
        "npl_qoq": ba.growth(npl, npl_prev_q),
        "cost_of_risk": _sd(pexp * ann, gross) if pexp is not None else None,
        "stress_base": _sd(E, A), "stress_x2": stress(2), "stress_x3": stress(3),
        # эх үүсвэр, хөрвөх чадвар
        "deposit_share": _sd(dep, L), "bank_borrow_share": _sd(bank_b, L), "other_src_share": _sd(oth, L),
        "rest_share": _sd(L - sum(x for x in (dep, bank_b, oth) if x is not None), L) if L and dep is not None else None,
        "casa": _sd(cur, cur + donly) if None not in (cur, donly) and donly else None,
        "liquid_assets": _sd(liquid, A), "liquid_to_deposits": _sd(liquid, dep),
        "ltd": m["ltd"],
        # ашгийн чанар
        "dupont": dup, "leverage": _sd(avgA, avgE),
        "roe_avg": _sd(np_ * ann, avgE) if np_ is not None else None,
        "ppop_assets": _sd(ppop, avgA), "ppop_loans": _sd(ppop, gross),
        "jaws": jaws,
        # капитал
        "icg": m["roe"], "asset_growth": m["assets_yoy"],
        "growth_gap": (m["assets_yoy"] - m["roe"]) if None not in (m["assets_yoy"], m["roe"]) else None,
        "tangible_equity": _sd(E - intang, A) if None not in (E, intang) else None,
    }


def justified_pb(adv: dict, vals: dict, metrics: dict, p: dict) -> dict:
    out = {}
    for b, v in vals.items():
        roe = metrics[b]["roe"]
        if roe is None or not v.get("pb"):
            continue
        jpb = (roe - p["growth"]) / (p["coe"] - p["growth"])
        out[b] = {"actual": v["pb"], "justified": jpb, "gap": v["pb"] / jpb - 1 if jpb > 0 else None}
    return out


def sector_hhi(data: dict[str, ba.BankData]) -> list[dict]:
    """Бүх банкны өгөгдөлтэй улирал бүрийн HHI (0–10000)."""
    periods = sorted(set.intersection(*[set(bd.bal) for bd in data.values()]))
    out = []
    for k in periods:
        row = {"period": f"{k[0]}Q{k[1]}"}
        for key, fn in (("assets", lambda bd: bd.b(k, "assets")), ("loans", lambda bd: bd.b(k, "loans_net")),
                        ("deposits", lambda bd: bd.deposits(k))):
            xs = [fn(bd) for bd in data.values()]
            if any(x is None for x in xs):
                row[key] = None
                continue
            t = sum(xs)
            row[key] = sum((x / t * 100) ** 2 for x in xs)
        out.append(row)
    return out


def ratio_history(data: dict[str, ba.BankData], fn, n: int = 10) -> dict[str, list]:
    """{банк: [(улирал, утга), …]} — сүүлийн n улирал."""
    out = {}
    for b, bd in data.items():
        k = bd.latest()
        ks, kk = [], k
        for _ in range(n):
            ks.insert(0, kk)
            kk = ba.prev_q(*kk)
        out[b] = [(f"{a}Q{c}", fn(bd, (a, c))) for a, c in ks]
    return out


def adv_highlights(adv: dict, jpb: dict, p: dict) -> list[str]:
    lines = []
    t = sorted([(b, a["texas"]) for b, a in adv.items() if a["texas"] is not None and a["texas"] >= p["texas_alert"]],
               key=lambda x: -x[1])
    if t:
        lines.append(f"Texas харьцаа {p['texas_alert']:.0%}-иас дээш (чанаргүй зээл / (ӨХ + эрсдэлийн сан)): "
                     + ", ".join(f"{b} ({x:.0%})" for b, x in t) + ".")
    h = [(b, a["overdue_npl_ratio"] - a["npl_ratio"]) for b, a in adv.items()
         if a["overdue_npl_ratio"] is not None and a["npl_ratio"] is not None]
    h = sorted([x for x in h if x[1] >= 0.05], key=lambda x: -x[1])
    if h:
        lines.append("Хугацаа хэтэрсэн зээл нь NPL-ийг 5+ нэгж хувиар нэмэгдүүлдэг (ирээдүйн NPL-ийн эх үүсвэр): "
                     + ", ".join(f"{b} (NPL {adv[b]['npl_ratio']:.1%} → {adv[b]['overdue_npl_ratio']:.1%})" for b, _ in h) + ".")
    g = sorted([(b, a["growth_gap"]) for b, a in adv.items() if a["growth_gap"] is not None and a["growth_gap"] >= p["gap_alert"]],
               key=lambda x: -x[1])
    if g:
        lines.append("Хөрөнгийн өсөлт дотоод капитал үүсгэлтээс (ROE) хол давсан — нэмэлт капитал татахгүй бол "
                     "ӨХ/хөрөнгийн харьцаа буурна: "
                     + ", ".join(f"{b} (өсөлт {adv[b]['asset_growth']:.0%} vs ROE {adv[b]['icg']:.0%})" for b, _ in g) + ".")
    s = sorted([(b, a["stress_x2"]) for b, a in adv.items() if a["stress_x2"] is not None], key=lambda x: x[1])[:2]
    if s:
        lines.append(f"Стресс тест (NPL ×2, алдагдал {p['lgd']:.0%}): ӨХ/хөрөнгө хамгийн их буурах — "
                     + ", ".join(f"{b} ({adv[b]['stress_base']:.1%} → {x:.1%})" for b, x in s) + ".")
    if jpb:
        lo = min(jpb.items(), key=lambda t: t[1]["gap"] if t[1]["gap"] is not None else 9)
        hi = max(jpb.items(), key=lambda t: t[1]["gap"] if t[1]["gap"] is not None else -9)
        lines.append(f"Үндэслэлтэй P/B-тэй харьцуулахад (COE {p['coe']:.0%}, g {p['growth']:.0%} таамаглалаар) хамгийн "
                     f"хямд: {lo[0]} (бодит {lo[1]['actual']:.2f}x vs {lo[1]['justified']:.2f}x), хамгийн үнэтэй: "
                     f"{hi[0]} ({hi[1]['actual']:.2f}x vs {hi[1]['justified']:.2f}x).")
    return lines


# ── Графикууд ─────────────────────────────────────────────────────────────
def _plt():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 9, "axes.facecolor": SURFACE, "figure.facecolor": SURFACE,
        "axes.edgecolor": GRID, "axes.labelcolor": INK2, "xtick.color": MUTED, "ytick.color": INK2,
        "axes.grid": True, "grid.color": GRID, "grid.linewidth": .8, "axes.axisbelow": True,
        "axes.spines.top": False, "axes.spines.right": False, "axes.titlesize": 11, "axes.titlecolor": INK,
        "axes.titleweight": "bold", "axes.titlelocation": "left", "legend.frameon": False,
        "legend.fontsize": 8.5,
    })
    return plt


def _pctfmt(ax, axis="x"):
    from matplotlib.ticker import MaxNLocator, PercentFormatter
    (ax.xaxis if axis == "x" else ax.yaxis).set_major_locator(MaxNLocator(nbins=6, steps=[1, 2, 2.5, 5, 10]))
    (ax.xaxis if axis == "x" else ax.yaxis).set_major_formatter(PercentFormatter(1.0, decimals=None))


def _save(fig, path, note=None):
    if note:
        fig.text(0.01, 0.01, note, fontsize=7.5, color=MUTED, ha="left", va="bottom")
    fig.savefig(path, dpi=160, facecolor=SURFACE)
    import matplotlib.pyplot as plt
    plt.close(fig)
    return path


def _label_points(ax, pts, fontsize=8):
    """Тархалтын цэгүүдийн нэрийг давхцуулахгүй байрлуулна: цэг бүрт хэд хэдэн
    байрлал туршиж, өмнө тавьсан шошготой (пикселийн ойролцоо хайрцгаар) давхцахгүйг сонгоно."""
    fig = ax.figure
    fig.canvas.draw()
    placed = []
    cands = [(6, 4, "left"), (6, -12, "left"), (-6, 4, "right"), (-6, -12, "right"),
             (6, 14, "left"), (6, -22, "left"), (-6, 14, "right"), (-6, -22, "right"), (0, 24, "center")]
    k = fig.dpi / 72.0                      # пункт → дэлгэцийн пиксель
    cw, ch = fontsize * .62 * k, fontsize * 1.3 * k
    for name, x, y in sorted(pts, key=lambda t: (-t[2], t[1])):
        px, py = ax.transData.transform((x, y))
        w = len(name) * cw
        for dx, dy, ha in cands:
            x0 = px + dx * k - (w if ha == "right" else w / 2 if ha == "center" else 0)
            box = (x0, py + dy * k, x0 + w, py + dy * k + ch)
            if not any(not (box[2] < b[0] or box[0] > b[2] or box[3] < b[1] or box[1] > b[3]) for b in placed):
                break
        placed.append(box)
        moved = (dx, dy) != (6, 4)
        ax.annotate(name, (x, y), xytext=(dx, dy), textcoords="offset points", ha=ha, fontsize=fontsize, color=INK2,
                    arrowprops={"arrowstyle": "-", "color": MUTED, "lw": .6, "shrinkA": 0, "shrinkB": 4} if moved else None)


def _hbars(ax, labels, series: list[tuple[str, list, str]], fmt="{:.1%}", ref=None, ref_label=None):
    """Хэвтээ бүлэглэсэн баганууд; 2px зай, утгын шошго, None → '—'."""
    import numpy as np
    n = len(series)
    h = .78 / n
    y = np.arange(len(labels))
    xmax = max([v for _, vals, _ in series for v in vals if v is not None] + [ref or 0])
    for i, (name, vals, color) in enumerate(series):
        pos = y - .39 + h * (i + .5)
        ax.barh(pos, [v if v is not None else 0 for v in vals], height=h * .88, color=color, label=name,
                edgecolor=SURFACE, linewidth=1)
        for yy, v in zip(pos, vals):
            ax.text((v or 0) + xmax * .01, yy, fmt.format(v) if v is not None else "—", va="center",
                    fontsize=7.5, color=INK2)
    ax.set_yticks(y)
    ax.set_yticklabels(labels)
    ax.invert_yaxis()
    ax.set_xlim(0, xmax * 1.18)
    ax.grid(axis="y", visible=False)
    if ref is not None:
        ax.axvline(ref, color=CRITICAL, lw=1, ls="--")
        if ref_label:
            ax.text(ref, len(labels) - .45, " " + ref_label, color=CRITICAL, fontsize=7.5, ha="left", va="bottom")


def make_advanced_charts(data, metrics, adv, jpb, hhi, out_dir: Path, p: dict) -> list[Path]:
    plt = _plt()
    import numpy as np
    out_dir = Path(out_dir)
    paths = []
    order = sorted(adv, key=lambda b: -(metrics[b]["assets"] or 0))
    dated = [f"{b} ({adv[b]['date'][2:]})" if adv[b]["date"] != max(a["date"] for a in adv.values()) else b for b in order]

    # 1. Зээлийн эрсдэл: NPL vs хугацаа хэтэрсэн+NPL | Texas
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(12, 5.2), gridspec_kw={"wspace": .5})
    _hbars(a1, dated, [("NPL", [adv[b]["npl_ratio"] for b in order], SERIES[0]),
                       ("Хугацаа хэтэрсэн + NPL", [adv[b]["overdue_npl_ratio"] for b in order], SERIES[1])])
    a1.set_title("NPL ба хугацаа хэтэрсэн зээл")
    _pctfmt(a1)
    a1.legend(loc="upper center", bbox_to_anchor=(.5, -.07), ncol=2)
    _hbars(a2, dated, [("Texas", [adv[b]["texas"] for b in order], SERIES[0])], fmt="{:.0%}",
           ref=p["texas_alert"], ref_label=f"{p['texas_alert']:.0%}")
    a2.set_title("Texas = NPL / (ӨХ + эрсдэлийн сан)")
    _pctfmt(a2)
    fig.subplots_adjust(left=.15, right=.98, top=.9, bottom=.17)
    paths.append(_save(fig, out_dir / "adv_credit_risk.png",
                       "'—' = хуудсанд хугацаа хэтэрсэн зээлийн мөр алга. Хаалтанд — өөр огнооны тайлан."))

    # 2. Стресс тест
    fig, ax = plt.subplots(figsize=(8, 4.8))
    _hbars(ax, dated, [("Одоо", [adv[b]["stress_base"] for b in order], SERIES[0]),
                       ("NPL ×2", [adv[b]["stress_x2"] for b in order], SERIES[1]),
                       ("NPL ×3", [adv[b]["stress_x3"] for b in order], SERIES[2])])
    ax.set_title("Стресс тест: өөрийн хөрөнгө / нийт хөрөнгө")
    _pctfmt(ax)
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.08), ncol=3)
    fig.tight_layout(rect=(0, .04, 1, 1))
    paths.append(_save(fig, out_dir / "adv_stress_test.png",
                       f"Шинэ чанаргүй зээлийн {p['lgd']:.0%} алдагдал болж ӨХ-ээс хасагдана гэж үзэв (татварын нөлөөгүй)."))

    # 3. Эх үүсвэрийн бүтэц — 100% давхарласан
    fig, ax = plt.subplots(figsize=(9, 4.8))
    parts = [("Харилцах ба хадгаламж", "deposit_share", SERIES[0]), ("Банк, СБ-аас татсан", "bank_borrow_share", SERIES[1]),
             ("Бусад эх үүсвэр", "other_src_share", SERIES[2]), ("Бусад өр төлбөр", "rest_share", NEUTRAL)]
    left = np.zeros(len(order))
    for name, key, color in parts:
        vals = np.array([adv[b][key] or 0 for b in order])
        ax.barh(dated, vals, left=left, color=color, label=name, edgecolor=SURFACE, linewidth=2, height=.62)
        for i, (l_, v) in enumerate(zip(left, vals)):
            if v >= .07:
                ax.text(l_ + v / 2, i, f"{v:.0%}", ha="center", va="center", fontsize=7.5,
                        color="#ffffff" if color in (SERIES[0], SERIES[1]) else INK)
        left += vals
    ax.invert_yaxis()
    ax.set_xlim(0, 1)
    _pctfmt(ax)
    ax.grid(axis="y", visible=False)
    ax.set_title("Эх үүсвэрийн бүтэц (нийт өр төлбөрт эзлэх хувь)")
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.08), ncol=4)
    fig.tight_layout()
    paths.append(_save(fig, out_dir / "adv_funding_mix.png"))

    # 4. Хөрвөх чадвар: хөрвөх хөрөнгө/хөрөнгө vs зээл/хадгаламж
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    for b in order:
        x, yv = adv[b]["liquid_assets"], adv[b]["ltd"]
        if None in (x, yv):
            continue
        ax.scatter(x, yv, s=60, color=SERIES[0], edgecolor=SURFACE, linewidth=2, zorder=3)
    ax.axhline(1, color=MUTED, lw=.8, ls="--")
    _pctfmt(ax, "x")
    _pctfmt(ax, "y")
    ax.set_xlabel("Хөрвөх хөрөнгө / нийт хөрөнгө  (мөнгө + банкинд байршуулсан + хөрөнгө оруулалт)")
    ax.set_ylabel("Цэвэр зээл / (харилцах + хадгаламж)")
    ax.set_title("Хөрвөх чадвар: баруун доод = илүү хөрвөх чадвартай")
    ax.margins(.15)
    fig.tight_layout()
    _label_points(ax, [(b, adv[b]["liquid_assets"], adv[b]["ltd"]) for b in order
                       if None not in (adv[b]["liquid_assets"], adv[b]["ltd"])])
    paths.append(_save(fig, out_dir / "adv_liquidity.png"))

    # 5. DuPont — ROA-гийн бүрэлдэхүүн (дундаж хөрөнгөд, жилийн хэмжээст)
    comp = [("Хүүгийн цэвэр орлого", "nii", 1, SERIES[0]), ("Хүүгийн бус орлого", "nonint_income", 1, SERIES[2]),
            ("Хүүгийн бус зардал", "opex", -1, SERIES[1]), ("Эрсдэлийн сангийн зардал", "provisions", -1, SERIES[3]),
            ("Татвар", "tax", -1, SERIES[4]), ("Бусад", "other", 1, NEUTRAL)]
    ob = [b for b in order if "other" in adv[b]["dupont"]]
    fig, ax = plt.subplots(figsize=(10, 5.2))
    pos_l, neg_l = np.zeros(len(ob)), np.zeros(len(ob))
    for name, key, sign, color in comp:
        vals = np.array([sign * adv[b]["dupont"][key] for b in ob])
        lefts = np.where(vals >= 0, pos_l, neg_l)
        ax.barh(range(len(ob)), vals, left=lefts, color=color, label=name, edgecolor=SURFACE, linewidth=2, height=.62)
        pos_l += np.where(vals >= 0, vals, 0)
        neg_l += np.where(vals < 0, vals, 0)
    roa = [adv[b]["dupont"]["roa"] for b in ob]
    ax.scatter(roa, range(len(ob)), marker="D", s=46, color=INK, zorder=4, label="ROA (цэвэр)")
    for i, r in enumerate(roa):
        ax.text(pos_l[i] + .001, i, f"ROA {r:.1%}", va="center", fontsize=7.5, color=INK2)
    ax.axvline(0, color=MUTED, lw=.8)
    ax.set_yticks(range(len(ob)))
    ax.set_yticklabels([dated[order.index(b)] for b in ob])
    ax.invert_yaxis()
    _pctfmt(ax)
    ax.grid(axis="y", visible=False)
    ax.set_title("DuPont: ROA хаанаас бүрддэг вэ (дундаж хөрөнгөд, жилийн хэмжээст)")
    ax.legend(loc="upper center", bbox_to_anchor=(.5, -.07), ncol=4)
    ax.set_xlim(neg_l.min() * 1.1, pos_l.max() * 1.25)
    fig.tight_layout()
    paths.append(_save(fig, out_dir / "adv_dupont.png"))

    # 6. Ашгийн дэр: PPOP/зээл vs зээлийн эрсдэлийн өртөг
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    pts = [(b, adv[b]["cost_of_risk"], adv[b]["ppop_loans"]) for b in order
           if None not in (adv[b]["cost_of_risk"], adv[b]["ppop_loans"])]
    for b, x, yv in pts:
        ax.scatter(x, yv, s=60, color=SERIES[0], edgecolor=SURFACE, linewidth=2, zorder=3)
    if pts:
        mx = max(max(x for _, x, _ in pts), max(y_ for _, _, y_ in pts)) * 1.1
        ax.plot([0, mx], [0, mx], color=CRITICAL, lw=1, ls="--")
        ax.text(.98, .04, "Тасархай шугам: PPOP = зээлийн алдагдал", transform=ax.transAxes, ha="right",
                color=CRITICAL, fontsize=7.5)
    _pctfmt(ax, "x")
    _pctfmt(ax, "y")
    ax.set_xlabel("Зээлийн эрсдэлийн өртөг (эрсдэлийн сангийн зардал / нийт зээл)")
    ax.set_ylabel("Сангийн өмнөх ашиг / нийт зээл")
    ax.set_title("Ашгийн дэр: шугамаас дээш хол байх тусам алдагдал даах чадвар их")
    ax.margins(.15)
    fig.tight_layout()
    _label_points(ax, pts)
    paths.append(_save(fig, out_dir / "adv_earnings_buffer.png"))

    # 7. Капиталын тогтвортой байдал: хөрөнгийн өсөлт vs ICG
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    pts = [(b, adv[b]["asset_growth"], adv[b]["icg"]) for b in order if None not in (adv[b]["asset_growth"], adv[b]["icg"])]
    for b, x, yv in pts:
        ax.scatter(x, yv, s=60, color=SERIES[1] if x - yv >= p["gap_alert"] else SERIES[0],
                   edgecolor=SURFACE, linewidth=2, zorder=3)
    if pts:
        mx = max(max(x for _, x, _ in pts), max(y_ for _, _, y_ in pts)) * 1.1
        ax.plot([0, mx], [0, mx], color=MUTED, lw=1, ls="--")
        ax.text(.98, .04, f"Улбар шар: өсөлт − ROE ≥ {p['gap_alert']:.0%}\n→ нэмэлт капиталгүйгээр ӨХ/хөрөнгө буурна",
                transform=ax.transAxes, ha="right", va="bottom", color=SERIES[1], fontsize=8)
        ax.text(.02, .96, "Шугамаас дээш: капитал өсөлтөөсөө\nтүрүүлж хуримтлагдана", transform=ax.transAxes,
                va="top", color=SERIES[0], fontsize=8)
    _pctfmt(ax, "x")
    _pctfmt(ax, "y")
    ax.set_xlabel("Нийт хөрөнгийн жилийн өсөлт")
    ax.set_ylabel("Дотоод капитал үүсгэлт (ROE, ногдол ашиггүй гэж)")
    ax.set_title("Өсөлтийн тогтвортой байдал")
    ax.margins(.12)
    fig.tight_layout()
    _label_points(ax, pts)
    paths.append(_save(fig, out_dir / "adv_capital_growth.png"))

    # 8. Үндэслэлтэй P/B
    if jpb:
        bs = sorted(jpb, key=lambda b: jpb[b]["actual"])
        fig, ax = plt.subplots(figsize=(8, 3.8))
        _hbars(ax, bs, [("Бодит P/B", [jpb[b]["actual"] for b in bs], SERIES[0]),
                        ("Үндэслэлтэй P/B", [jpb[b]["justified"] for b in bs], SERIES[1])], fmt="{:.2f}x", ref=1.0)
        ax.set_title("Бодит ба үндэслэлтэй P/B = (ROE − g) / (COE − g)")
        ax.legend(loc="upper center", bbox_to_anchor=(.5, -.1), ncol=2)
        fig.tight_layout(rect=(0, .05, 1, 1))
        paths.append(_save(fig, out_dir / "adv_justified_pb.png",
                           f"Таамаглал: COE {p['coe']:.0%}, g {p['growth']:.0%} (--coe, --growth-оор өөрчилнө). "
                           "ROE — тайлант үеийн жилийн хэмжээст."))

    # 9. Салбарын төвлөрөл (HHI)
    rows = [r for r in hhi if any(r[k] is not None for k in ("assets", "loans", "deposits"))]
    if len(rows) >= 2:
        fig, ax = plt.subplots(figsize=(8, 4))
        xs = [r["period"] for r in rows]
        for (name, key), color in zip((("Хөрөнгө", "assets"), ("Зээл", "loans"), ("Харилцах+хадгаламж", "deposits")), SERIES):
            ys = [r[key] for r in rows]
            ax.plot(xs, ys, color=color, lw=2, marker="o", ms=4, label=name)
            ax.annotate(f"{ys[-1]:,.0f}", (len(xs) - 1, ys[-1]), xytext=(6, 0), textcoords="offset points",
                        fontsize=7.5, color=INK2, va="center")
        ax.axhline(1800, color=MUTED, lw=.8, ls="--")
        ax.text(0, 1800, " 1,800 — өндөр төвлөрлийн уламжлалт босго", fontsize=7.5, color=MUTED, va="bottom")
        ax.set_title("Салбарын төвлөрөл (HHI, 9 банк бүгд тайлантай улирлууд)")
        ax.legend(loc="upper center", bbox_to_anchor=(.5, -.2), ncol=3)
        ax.tick_params(axis="x", rotation=45)
        fig.tight_layout()
        paths.append(_save(fig, out_dir / "adv_sector_hhi.png"))

    # 10–11. Чиг хандлага — банк бүр тусдаа самбар (small multiples), салбарын медиан саарал
    for key, title, fname in (("npl_ratio", "Чанаргүй зээлийн харьцаа", "adv_trend_npl.png"),
                              ("roe", "ROE (жилийн хэмжээст)", "adv_trend_roe.png")):
        def fn(bd, k, key=key):
            if not bd.bal.get(k) or not bd.inc.get(k):
                return None
            try:
                return ba.compute_bank(bd, k)[key]
            except Exception:  # noqa: BLE001
                return None
        hist = ratio_history(data, fn, n=10)
        periods = sorted({pp for v in hist.values() for pp, _ in v})
        med = {pp: statistics.median([x for b in hist for q_, x in hist[b] if q_ == pp and x is not None] or [0])
               for pp in periods}
        fig, axes = plt.subplots(3, 3, figsize=(11, 7.5), sharex=False, sharey=True)
        for ax, b in zip(axes.flat, order):
            ser = hist[b]
            xs_ = [periods.index(pp) for pp, x in ser if x is not None]
            ys_ = [x for _, x in ser if x is not None]
            ax.plot(range(len(periods)), [med[pp] for pp in periods], color=NEUTRAL, lw=1.5, label="Салбарын медиан")
            ax.plot(xs_, ys_, color=SERIES[0], lw=2, marker="o", ms=3, label=b)
            if ys_:
                ax.annotate(f"{ys_[-1]:.1%}", (xs_[-1], ys_[-1]), xytext=(4, 4), textcoords="offset points",
                            fontsize=7.5, color=INK2)
            ax.set_title(b, fontsize=9.5)
            ax.set_xticks([0, len(periods) - 1])
            ax.set_xticklabels([periods[0], periods[-1]], fontsize=7.5)
            _pctfmt(ax, "y")
        for ax in list(axes.flat)[len(order):]:
            ax.set_visible(False)
        axes.flat[0].legend(loc="upper left", fontsize=7)
        fig.suptitle(f"{title} — банк бүр (цэнхэр) ба салбарын медиан (саарал)", x=.01, ha="left",
                     fontsize=11, fontweight="bold", color=INK)
        fig.tight_layout(rect=(0, 0, 1, .96))
        paths.append(_save(fig, out_dir / fname))
    return paths


# ── Word хүснэгт ──────────────────────────────────────────────────────────
def table_rows(adv: dict, order: list[str]) -> list[tuple[str, list[str], list[list[str]]]]:
    fp = ba.fp

    def x(v):
        return "—" if v is None else f"{v:.2f}x"
    return [
        ("Зээлийн эрсдэл", ["Банк", "NPL", "Хуг.хэтэрсэн+NPL", "Texas", "Хоцрогдсон NPL", "CoR", "ӨХ/Хөр: одоо", "NPL×2", "NPL×3"],
         [[b, fp(adv[b]["npl_ratio"]), fp(adv[b]["overdue_npl_ratio"]), fp(adv[b]["texas"]), fp(adv[b]["npl_lagged"]),
           fp(adv[b]["cost_of_risk"]), fp(adv[b]["stress_base"]), fp(adv[b]["stress_x2"]), fp(adv[b]["stress_x3"])]
          for b in order]),
        ("Эх үүсвэр ба хөрвөх чадвар", ["Банк", "Харилцах+хадг./өр", "Банк, СБ/өр", "Харилцахын хувь",
                                        "Хөрвөх хөр./хөр.", "Хөрвөх хөр./хадг.", "Зээл/хадг."],
         [[b, fp(adv[b]["deposit_share"]), fp(adv[b]["bank_borrow_share"]), fp(adv[b]["casa"]),
           fp(adv[b]["liquid_assets"]), fp(adv[b]["liquid_to_deposits"]), fp(adv[b]["ltd"])] for b in order]),
        ("Ашгийн чанар ба капитал", ["Банк", "ROA (дундаж)", "Хөшүүрэг", "ROE (дундаж)", "PPOP/хөр.", "Jaws",
                                     "Хөр. өсөлт", "ICG", "Биет ӨХ/хөр."],
         [[b, fp(adv[b]["dupont"].get("roa")), x(adv[b]["leverage"]), fp(adv[b]["roe_avg"]), fp(adv[b]["ppop_assets"]),
           fp(adv[b]["jaws"], True), fp(adv[b]["asset_growth"], True), fp(adv[b]["icg"]), fp(adv[b]["tangible_equity"])]
          for b in order]),
    ]


DEFINITIONS = [
    "Хугацаа хэтэрсэн + NPL = (хугацаа хэтэрсэн/анхаарал хандуулах + чанаргүй зээл) / нийт зээл.",
    "Texas харьцаа = чанаргүй зээл / (өөрийн хөрөнгө + зээлийн эрсдэлийн сан); 100%-д ойртох тусам дэр нимгэн.",
    "Хоцрогдсон NPL = одоогийн чанаргүй зээл / жилийн өмнөх нийт зээл (хурдан өсөлтийн 'шингэрүүлэлт'-ийг арилгана).",
    "CoR = жилийн хэмжээст эрсдэлийн сангийн зардал / нийт зээл.",
    "Стресс тест: NPL 2 ба 3 дахин өсөж, шинэ чанаргүй зээлийн LGD хувь нь алдагдал болж өөрийн хөрөнгөөс хасагдана.",
    "Хөрвөх хөрөнгө = мөнгө + банк, СБ-д байршуулсан + хөрөнгө оруулалт. Харилцахын хувь = харилцах / (харилцах + хадгаламж).",
    "DuPont: ROA = (ЦХО + хүүгийн бус орлого − хүүгийн бус зардал − эрсдэлийн сан − татвар ± бусад) / дундаж хөрөнгө; "
    "ROE (дундаж) = ROA × хөшүүрэг (дундаж хөрөнгө / дундаж ӨХ). Дундаж = оны эхний ба тайлант үеийн үлдэгдлийн дундаж.",
    "PPOP = сангийн өмнөх ашиг = ЦХО + хүүгийн бус орлого − хүүгийн бус зардал. Jaws = орлогын өсөлт − зардлын өсөлт (YoY).",
    "ICG = дотоод капитал үүсгэлт = ROE × (1 − ногдол ашгийн хувь); ногдол ашгийн мэдээлэлгүй тул ROE-тэй тэнцүү гэж үзэв.",
    "Үндэслэлтэй P/B = (ROE − g) / (COE − g); COE, g нь таамаглал.",
    "HHI = банкуудын зах зээлийн хувийн квадратын нийлбэр (0–10,000), зөвхөн 9 банк бүгд өгөгдөлтэй улирлаар.",
]
