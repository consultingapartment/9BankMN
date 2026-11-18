#!/usr/bin/env python3
"""
test_extended_analysis.py — өргөтгөсөн шинжилгээний томьёог шалгана.
    python3 test_extended_analysis.py <Q3_workbook.xlsx>
"""
import math
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import advanced_ratios as ar  # noqa: E402
import bank_analysis as ba  # noqa: E402
import extended_analysis as ea  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


def main(wbp):
    wb = load_workbook(wbp, data_only=True)
    data = {b: ba.BankData(wb, b) for b in ba.BANKS}
    ar.attach_extra(wb, data)
    metrics = {b: ba.compute_bank(bd, bd.latest()) for b, bd in data.items()}
    adv = {b: ar.compute_advanced(data[b], data[b].latest(), metrics[b], ar.DEFAULTS) for b in data}

    print("═══ Транс банк 2026/09/30 — эх тайлангийн тоогоор гараар ═══")
    bd, k = data["Транс банк"], (2026, 3)
    blk = ea.bank_block(bd, k, metrics["Транс банк"])
    iea_now = 221948 + 405285 + 1064386
    iea0 = sum(ar._extra(bd, (2025, 4), x) for x in ("placements", "investments")) + bd.b((2025, 4), "loans_net")
    check(abs(blk["asset_yield"] - 194260 * 4 / 3 / ((iea_now + iea0) / 2)) < 1e-9, f"активын өгөөж = {blk['asset_yield']:.2%}")
    ibl_now = 153461 + 1189011 + 165833 + 419852
    ibl0 = bd.deposits((2025, 4)) + ar._extra(bd, (2025, 4), "bank_borrow") + ar._extra(bd, (2025, 4), "other_src")
    check(abs(blk["funding_cost"] - 149081 * 4 / 3 / ((ibl_now + ibl0) / 2)) < 1e-9, f"эх үүсвэрийн өртөг = {blk['funding_cost']:.2%}")
    check(abs(blk["eff_tax"] - 3137 / 11523) < 1e-9, f"татварын бодит хувь = {3137 / 11523:.1%}")
    check(abs(sum(blk[x] for x in ("mix_cash", "mix_placements", "mix_investments", "mix_loans", "mix_other")) - 1) < 1e-9,
          "активын бүтэц нийлбэр = 100%")
    check(abs(blk["reserve_ratio"] - 59097 / (1064386 + 59097)) < 1e-9, "нөөцийн харьцаа = сан / нийт зээл")

    print("\n═══ Мэдээлэл дутуу үед '—' ═══")
    check(ea.bank_block(data["Капитрон"], data["Капитрон"].latest(), metrics["Капитрон"])["asset_yield"] is None,
          "Капитрон: байршуулсан мөргүй → активын өгөөж None (хэтрүүлэхгүй)")
    check(ea.bank_block(data["ХХБ"], data["ХХБ"].latest(), metrics["ХХБ"])["eff_tax"] is None,
          "ХХБ: ТӨА ≈ цэвэр ашиг → татварын хувь None (0% биш)")

    print("\n═══ Зах зээлийн хувь, скоркард ═══")
    sh = ea.market_share(data)
    for key in ("assets", "loans", "deposits"):
        check(abs(sum(v[key] for v in sh["banks"].values()) - 1) < 1e-9, f"{key} хувийн нийлбэр = 100% ({sh['period']})")
    check(abs(sum(v["assets_chg"] for v in sh["banks"].values())) < 1e-9, "хувийн өөрчлөлтийн нийлбэр = 0")
    sc = ea.scorecard(metrics, adv)
    for name, *_ in ea.SCORE_METRICS:
        rs = sorted(r["ranks"][name] for r in sc.values() if r["ranks"][name] is not None)
        if rs != list(range(1, len(rs) + 1)):
            check(False, f"{name}: эрэмбэ 1..n биш {rs}")
    check(sorted(v["overall"] for v in sc.values()) == list(range(1, 10)), "нийлмэл эрэмбэ 1..9")
    best = min(sc, key=lambda b: sc[b]["overall"])
    check(all(sc[best]["score"] >= v["score"] for v in sc.values()), f"1-р байр хамгийн өндөр оноотой ({best})")

    print("\n═══ Хувьцааны статистик (зохиомол ханш) ═══")
    d0 = date(2025, 9, 1)
    dates = [(d0 + timedelta(days=i)).isoformat() for i in range(366)]
    close = [100 * (1.001 ** i) for i in range(366)]
    st = ea.stock_stats({"Х": {"date": dates, "close": close, "volume": [1000] * 366}},
                        {"Х": {"shares": 1_000_000}}, {"Х": {"ff_pct": .1}})["Х"]
    check(abs(st["r_1y"] - (close[-1] / close[1] - 1)) < 1e-9 or abs(st["r_1y"] - (1.001 ** 365 - 1)) < 1e-6, f"1 жилийн өгөөж {st['r_1y']:.2%}")
    check(st["max_dd_1y"] == 0 and abs(st["from_high"]) < 1e-12, "тасралтгүй өсөлт → уналт 0, оргил дээр")
    check(abs(st["vol_1y"]) < 1e-6, "тогтмол өсөлт → хэлбэлзэл ≈ 0")
    check(abs(st["turnover_1y"] - 1000 * 365 / 100_000) < 1e-9, "эргэлт = жилийн арилжсан / чөлөөт эргэлтийн хувьцаа")
    sparse = dates[::3]   # 3 хоног тутамд нэг арилжаа
    st2 = ea.stock_stats({"Х": {"date": sparse, "close": close[::3], "volume": [1000] * len(sparse)}},
                         {"Х": {"shares": 1_000_000}}, {})["Х"]
    check(.4 < st2["days_traded_3m"] < .55, f"сийрэг арилжаа → арилжаатай өдөр {st2['days_traded_3m']:.0%} (100% биш)")

    print("\n═══ Прогноз ба Graham ═══")
    class Fake:
        def quarter_alone(self, k, key):
            return 100 + 10 * (k[0] * 4 + k[1])
        def ytd(self, k, key):
            return sum(self.quarter_alone((k[0], i), key) for i in range(1, k[1] + 1))
    f = ea.forecast(Fake(), (2026, 2))
    exp = 100 + 10 * (2026 * 4 + 3)
    check(abs(f["forecast"] - exp) < 1e-6 and abs(f["high"] - f["low"]) < 1e-6, f"шугаман цуваа → яг {exp} (муж 0)")
    check(f["next"] == "2026Q3" and abs(f["fy_estimate"] - Fake().ytd((2026, 4), "")) < 1e-6, "жилийн тооцоо = YTD + үлдсэн улирлууд")
    v = ea.valuation_extra({"Т": {"eps": 100, "bvps": 1000, "price": 500, "pe": 5, "pb": .5, "mcap": 50}},
                           {"Т": {"roe": .1, "assets": 1000, "profit_ytd_yoy": .25}})["Т"]
    check(abs(v["graham"] - math.sqrt(22.5 * 100 * 1000)) < 1e-9 and abs(v["peg"] - .2) < 1e-12, "Graham = √(22.5·EPS·BVPS), PEG = P/E / (g×100)")

    print("\n═══ Графикууд ═══")
    ext = ea.compute(data, metrics, adv, {}, {}, {})
    ch = ea.make_charts(data, metrics, adv, ext, Path(tempfile.mkdtemp()))
    n = sum(len(v) for v in ch.values())
    check(n >= 9 and all(p.stat().st_size > 15000 for v in ch.values() for p in v), f"{n} график (ханш/үнэлгээгүй үед)")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
