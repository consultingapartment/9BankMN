#!/usr/bin/env python3
"""
test_advanced_ratios.py — нэмэлт харьцаануудын томьёог гараар тооцсон утгатай тулгана.
    python3 test_advanced_ratios.py <Q3_workbook.xlsx>
"""
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import advanced_ratios as ar  # noqa: E402
import bank_analysis as ba  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


def main(wbp):
    wb = load_workbook(wbp, data_only=True)
    data = {b: ba.BankData(wb, b) for b in ba.BANKS}
    ar.attach_extra(wb, data)
    p = dict(ar.DEFAULTS)
    metrics = {b: ba.compute_bank(bd, bd.latest()) for b, bd in data.items()}
    adv = {b: ar.compute_advanced(data[b], data[b].latest(), metrics[b], p) for b in data}

    print("═══ Богд банк 2026/09/30 (эх тайлангийн сая ₮ утгаар гараар) ═══")
    a = adv["Богд банк"]
    net, prov, npl, od, eq, A = 698932.14949, 7067.67191, 14359.4776, 58975.97318, 237000.7157, 1451853.14317
    gross = net + prov
    check(abs(a["overdue_npl_ratio"] - (npl + od) / gross) < 1e-9, f"хуг.хэтэрсэн+NPL = {(npl + od) / gross:.2%}")
    check(abs(a["texas"] - npl / (eq + prov)) < 1e-9, f"Texas = {npl / (eq + prov):.2%}")
    check(abs(a["stress_x2"] - (eq - npl * .5) / (A - npl * .5)) < 1e-9, "стресс ×2 = (ӨХ − NPL×LGD)/(Хөр − NPL×LGD)")
    check(abs(a["stress_x3"] - (eq - 2 * npl * .5) / (A - 2 * npl * .5)) < 1e-9, "стресс ×3")
    check(a["casa"] is None, "Богд: харилцах/хадгаламж нэгдсэн мөр → CASA '—' (тэг биш)")
    check(abs(a["deposit_share"] - 483902.25064 / 1214852.42747) < 1e-9, "харилцах+хадгаламж / нийт өр")

    print("\n═══ Ерөнхий шинж ═══")
    for b, x in adv.items():
        d = x["dupont"]
        if "other" in d:
            s = d["nii"] + d["nonint_income"] - d["opex"] - d["provisions"] - d["tax"] + d["other"]
            check(abs(s - d["roa"]) < 1e-12, f"{b}: DuPont бүрэлдэхүүний нийлбэр = ROA ({d['roa']:.2%})")
        if x["roe_avg"] is not None and x["leverage"] is not None:
            check(abs(x["roe_avg"] - d["roa"] * x["leverage"]) < 1e-9, f"{b}: ROE(дундаж) = ROA × хөшүүрэг")
    check(adv["Хаан банк"]["overdue_npl_ratio"] is None and adv["Голомт"]["overdue_npl_ratio"] is None,
          "Хаан, Голомт: хугацаа хэтэрсэн мөргүй → None")
    hhi = ar.sector_hhi(data)
    check(all(1000 < r["assets"] < 10000 for r in hhi if r["assets"]), f"HHI хүрээнд ({hhi[-1]['assets']:,.0f})")
    t = {"Төрийн банк": {"pb": .48}}
    j = ar.justified_pb(adv, t, metrics, p)["Төрийн банк"]
    exp = (metrics["Төрийн банк"]["roe"] - .06) / (.20 - .06)
    check(abs(j["justified"] - exp) < 1e-12, f"үндэслэлтэй P/B = (ROE − g)/(COE − g) = {exp:.2f}")
    out = Path(tempfile.mkdtemp())
    paths = ar.make_advanced_charts(data, metrics, adv, ar.justified_pb(adv, {"Төрийн банк": {"pb": .48}}, metrics, p),
                                    hhi, out, p)
    check(len(paths) == 11 and all(x.stat().st_size > 20000 for x in paths), f"{len(paths)} график үүссэн")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
