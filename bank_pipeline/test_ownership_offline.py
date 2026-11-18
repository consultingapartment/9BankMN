#!/usr/bin/env python3
"""
test_ownership_offline.py — ownership_source.py + шинжилгээний залгаасыг шалгана.
    python3 test_ownership_offline.py <uctht_concentration.xlsx> <Q3_workbook.xlsx> <price_data.json> <fundamentals.json>
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_analysis as ba  # noqa: E402
import ownership_source as os_  # noqa: E402
import quarterly_pipeline as qp  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


def synthetic(path: Path, sheets: dict):
    """Баганын дараалал ӨӨР (чөлөөт эргэлтийн блок түрүүлж, нэмэлт багана) файл."""
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        ws.append(["ҮНЭТ ЦААСНЫ ТӨВЛӨРЛИЙН СУДАЛГАА"])
        ws.append([])
        ws.append(["№", "Үнэт цаасны нэр", "ҮЦ код", "Тэмдэглэл", "5-аас бага хувь", None, None,
                   "Нийт гаргасан\nҮЦ-ны тоо", "5-аас их хувь", None, None, "Төрийн мэдэлд", None])
        ws.append([None, None, None, None, "Эзэмшигч\n-дийн тоо", "Эзлэх\nхувь", "ҮЦ-ны тоо\nширхэг", None,
                   "Эзлэх\nхувь", "Эзэмшигч\n-дийн тоо", "ҮЦ-ны тоо\nширхэг", "Эзлэх\nхувь", "ҮЦ-ны тоо\nширхэг"])
        for i, (code, name, total, dom, domh, ffh) in enumerate(rows, 1):
            ws.append([i, name, code, "x", ffh, round(1 - dom, 4), total * (1 - dom), total, dom, domh, total * dom, 0, 0])
    wb.save(path)


def main(real_xlsx, q3_wb, prices, fund):
    work = Path(tempfile.mkdtemp(prefix="own_test_"))

    print("═══ 1. Бодит файл ═══")
    snaps = os_.load_snapshots(Path(real_xlsx))
    sbm = snaps["Төрийн банк"][-1]
    check(len(snaps["Хаан банк"]) >= 17, f"Хаан банк {len(snaps['Хаан банк'])} тайлан")
    check(sbm["total"] == 755096250 and abs(sbm["ff_pct"] - 0.0571) < 1e-9, "Төрийн банк: хувьцаа 755,096,250, чөлөөт 5.71%")
    summ = os_.summarize(snaps, "2026-09-30")
    check(any("82.4% → 85.0%" in e["text"] for e in summ["Голомт"]["events"]), "Голомт 2025-12: давамгай хувь 82.4% → 85.0% үйл явдал")
    x = summ["ХХБ"]
    check(x["under5"]["holders_3m"] > 0.4, f"ХХБ 5%-иас ДООШ эзэмшигчдийн 3 сарын өсөлт {x['under5']['holders_3m']:+.1%}")
    check(x["over5"]["holders"] == 2 and x["under5"]["holders"] == 4786, "ХХБ: 5%-иас ДЭЭШ 2, ДООШ 4,786 — ТУСДАА (нийлбэр биш)")
    check(abs(x["over5"]["pct"] + x["under5"]["pct"] + x["state_pct"] - 1) < 1e-3, "бүлгүүдийн эзлэх хувийн нийлбэр = 100%")
    check(x["over5"]["shares"] + x["under5"]["shares"] == x["total_shares"], "бүлгүүдийн хувьцааны нийлбэр = гаргасан хувьцаа")
    import openpyxl
    sm = openpyxl.load_workbook(real_xlsx, data_only=True)["Эзэмшигчдийн тоо"]
    src = {sm.cell(r, 2).value: sm.cell(r, sm.max_column - 1).value for r in range(5, sm.max_row + 1)}
    check(all(src[n] == summ[b]["under5"]["holders"] for b, n in (("Хаан банк", "ХААН БАНК"), ("Төрийн банк", "ТӨРИЙН БАНК"),
                                                                    ("ХХБ", "ХУДАЛДАА ХӨГЖЛИЙН БАНК"))),
          "5%-иас ДООШ эзэмшигчдийн тоо эх файлын 'Эзэмшигчдийн тоо' хуудастай яг тэнцүү")
    g = summ["Голомт"]
    check(any(e["group"] == "over5" and "82.4% → 85.0%" in e["text"] for e in g["events"]), "Голомт: 5%-иас ДЭЭШ бүлгийн өөрчлөлт тусдаа үйл явдал")
    check(summ["Хаан банк"]["qend_as_of"] == "2026-09-30", "улирлын эцсийн snapshot = 2026-09-30")

    print("\n═══ 2. Өөр баганын дараалал + хувьцааны тоо өөрчлөгдсөн ═══")
    syn = work / "syn.xlsx"
    synthetic(syn, {
        "Тайлан 26-06-30": [(564, "ТӨРИЙН БАНК", 755096250, 0.9429, 2, 12919), (563, "ХААН БАНК", 1912198000, 0.8577, 4, 40657)],
        "Тайлан 26-09-30": [(564, "ТӨРИЙН БАНК", 755096250, 0.9429, 2, 13549), (563, "ХААН БАНК", 1912198000, 0.8577, 4, 43499)],
        "Тайлан 26-10-31": [(564, "ТӨРИЙН БАНК", 830605875, 0.90, 2, 16000), (563, "ХААН БАНК", 1912198000, 0.8577, 4, 45092)],
        "Хураангуй": [],
    })
    s2 = os_.summarize(os_.load_snapshots(syn))
    t = s2["Төрийн банк"]
    check(abs(t["ff_pct"] - 0.10) < 1e-9 and t["under5"]["holders"] == 16000 and t["over5"]["holders"] == 2,
          "баганыг гарчгаар нь олсон (дараалал хамаарахгүй)")
    kinds = {e["type"] for e in t["events"]}
    check({"shares", "over5_pct", "under5_surge"} <= kinds, f"үйл явдлууд: {sorted(kinds)}")
    metrics = {"Төрийн банк": {"profit_ttm": 140000, "equity": 750000, "date": "2026/09/30"}}
    pr = {"Төрийн банк": {"date": ["2026-10-01"], "close": [500]}}
    v = ba.valuation(metrics, pr, {"Төрийн банк": 755096250}, s2)["Төрийн банк"]
    check(v["shares"] == 830605875 and v["shares_note"], f"EPS шинэ хувьцааны тоогоор + анхааруулга ({v['shares_note']})")
    check(abs(v["ff_mcap"] - 500 * 830605875 * 0.10 / 1e6) < 1e-6, "чөлөөт эргэлтийн үнэлгээ")

    print("\n═══ 3. Татаж чадахгүй үед ═══")
    cache = work / "cache" / "u.xlsx"
    cache.parent.mkdir()
    shutil.copy(real_xlsx, cache)
    s3, note = os_.load("http://127.0.0.1:9/none.xlsx", cache)
    check(bool(s3) and "кэш" in note, f"кэшийг ашигласан: {note[:60]}")
    s4, note4 = os_.load("http://127.0.0.1:9/none.xlsx", work / "nocache" / "u.xlsx")
    check(s4 == {} and "ороогүй" in note4, "кэш ч алга → эзэмшлийн хэсэггүй үргэлжилнэ")
    f2 = work / "f2.json"
    shutil.copy(fund, f2)
    r = ba.run(q3_wb, str(work / "an0"), prices, str(f2), ownership={}, ownership_note=note4)
    check(r["docx"].exists() and "Эзэмшлийн бүтэц" not in json.dumps(r["highlights"], ensure_ascii=False),
          "эзэмшлийн мэдээлэлгүй ч тайлан гарсан")

    print("\n═══ 4. Pipeline --analysis-only --ownership-file ═══")
    f3 = work / "f3.json"
    shutil.copy(fund, f3)
    rc = qp.main(["--workbook", q3_wb, "--analysis-only", "--analysis-dir", str(work / "an1"),
                  "--prices", prices, "--fundamentals", str(f3), "--ownership-file", real_xlsx])
    check(rc == 0, "exit 0")
    d = json.loads((work / "an1" / "2026Q3" / "metrics.json").read_text(encoding="utf-8"))
    check(set(d["ownership"]) == set(ba.MSE_BANKS), "5 банкны эзэмшлийн мэдээлэл metrics.json-д")
    check(all(d["valuation"][b].get("ff_pct") is not None for b in ba.MSE_BANKS), "valuation-д чөлөөт эргэлт")
    fj = json.loads(f3.read_text(encoding="utf-8"))
    check(all(isinstance(fj[b].get("holders_over5"), int) and isinstance(fj[b].get("holders_under5"), int)
              and isinstance(fj[b]["shares"], int) and "holders" not in fj[b] for b in ba.MSE_BANKS),
          "fundamentals.json: holders_over5 / holders_under5 тусдаа, нийлбэр 'holders' байхгүй")
    tg = (work / "an1" / "2026Q3" / "telegram.txt").read_text(encoding="utf-8")
    check("5%-иас ДООШ" in tg and "5%-иас ДЭЭШ" in tg, "дүгнэлтэд хоёр бүлэг тусдаа")
    from docx import Document as _D
    dtx = "\n".join(p.text for p in _D(next((work / "an1" / "2026Q3").glob("*.docx"))).paragraphs)
    check("5%-иас ДЭЭШ эзэмшигчид" in dtx and "5%-иас ДООШ эзэмшигчид (олон нийт, чөлөөт эргэлт)" in dtx,
          "Word: 5%-иас дээш ба доош эзэмшигчдийн тусдаа хүснэгт")
    check((work / "an1" / "2026Q3" / "own_structure.png").exists() and (work / "an1" / "2026Q3" / "own_history.png").exists(),
          "эзэмшлийн бүтэц ба түүхийн графикууд")
    check((work / "an1" / "2026Q3" / "chart_pb_roe_freefloat.png").exists(), "P/B–ROE–чөлөөт эргэлтийн график")
    from docx import Document
    txt = "\n".join(p.text for p in Document(next((work / "an1" / "2026Q3").glob("*.docx"))).paragraphs)
    import re as _re
    heads = [p.text for p in Document(next((work / "an1" / "2026Q3").glob("*.docx"))).paragraphs
             if p.style.name == "Heading 1"]
    nums = [int(_re.match(r"(\d+)\.", h).group(1)) for h in heads if _re.match(r"\d+\.", h)]
    check("5. Эзэмшлийн бүтэц ба хувьцааны хөрвөх чадвар (МХБ-ийн 5 банк)" in heads and nums == list(range(1, len(nums) + 1))
          and heads[-1].endswith("Графикууд") and len(heads) >= 18,
          f"Word: {len(heads)} бүлэг, 5-р нь эзэмшил, дугаарлалт 1..{len(nums)} тасралтгүй")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:5]))
