#!/usr/bin/env python3
"""
test_bogdbank_offline.py — bogdbank_source.py + pipeline-ийг сүлжээгүй шалгана.
    python3 test_bogdbank_offline.py <Q3_workbook.xlsx>
"""
import json
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402
import bogdbank_source as bs  # noqa: E402
import capitron_source as cs  # noqa: E402
import quarterly_pipeline as qp  # noqa: E402
import transbank_source as ts  # noqa: E402
from test_capitron_offline import Resp  # noqa: E402
from test_transbank_offline import INC, STD  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


def rows_from(ref, date_, bm, im):
    wb, wi = ref["Богд банк баланс"], ref["Богд банк орлого"]
    cb, ci = qp.find_date_col(wb, 4, date_), qp.find_date_col(wi, 4, date_)
    out = []
    for ws, labs, c, m in ((wb, STD + ["Хугацаа хэтэрсэн зээл"], cb, bm), (wi, INC, ci, im)):
        for lab in labs:
            r = next((rr for rr in range(5, ws.max_row + 1) if str(ws.cell(rr, 1).value or "").strip() == lab), None)
            v = ws.cell(r, c).value if r else None
            if isinstance(v, (int, float)):
                if lab == "Чанаргүй зээл":      # сайт 3 ангиллаар задалж өгдөг гэж үзье
                    out += [("Хэвийн бус зээл", v * .2 * m), ("Эргэлзээтэй зээл", v * .3 * m), ("Муу зээл", v * .5 * m)]
                else:
                    out.append((lab, v * m))
    return out


def html_table(rows):
    """Мянган ₮, '1 234 567.8' зайтай тоо, (…) сөрөг, '1.1' дугаар — бодит сайттай төстэй."""
    def f(x):
        s = f"{abs(x) * 1000:,.1f}".replace(",", " ")
        return f"({s})" if x < 0 else s
    trs = "".join(f"<tr><td>{i}.</td><td>{lab}</td><td>{f(v * .95)}</td><td>{f(v)}</td></tr>"
                  for i, (lab, v) in enumerate(rows, 1))
    return (f'<div class="report"><p class="text-right">/Мянган төгрөг/</p><table class="table">'
            f'<thead><tr><th>№</th><th>Үзүүлэлт</th><th>Эхний үлдэгдэл</th><th>Эцсийн үлдэгдэл</th></tr></thead>'
            f'<tbody><tr><td></td><td>ХӨРӨНГӨ</td><td></td><td></td></tr>{trs}</tbody></table></div>')


class FakeSession:
    def __init__(self, pages):
        self.pages, self.calls = pages, []

    def post(self, url, data=None, headers=None, timeout=60):
        self.calls.append((data["year"], data["session"], data["report_type"]))
        r = Resp()
        r.text = self.pages.get((data["year"], data["session"]), '<div class="report"><p>Мэдээлэл олдсонгүй</p></div>')
        return r


def main(q3_wb):
    work = Path(tempfile.mkdtemp(prefix="bogd_test_"))
    ref = load_workbook(q3_wb, data_only=True)
    q3 = rows_from(ref, "2026/06/30", 1.03, 1.5)

    print("═══ 1. HTML задлалт ═══")
    raw = bs.parse_html(html_table(q3))
    exp = dict(q3)
    check(abs(raw["НИЙТ ХӨРӨНГИЙН ДҮН"] - exp["НИЙТ ХӨРӨНГИЙН ДҮН"] * 1000) < 1, "'1 234 567.8' зайтай тоо, баруун багана")
    check(raw["Зээлийн эрсдэлийн сан"] < 0, "(…) → сөрөг")
    check(abs(raw["Чанаргүй зээл"] - (exp["Хэвийн бус зээл"] + exp["Эргэлзээтэй зээл"] + exp["Муу зээл"]) * 1000) < 1,
          "Хэвийн бус + Эргэлзээтэй + Муу → Чанаргүй зээл")
    js = json.dumps({"status": True, "html": html_table(q3)})
    check(len(bs.parse_html(js)) == len(raw), "JSON {html: …} хэлбэрийн хариу ч задлагдана")

    print("\n═══ 2. Улирал асуух (workbook 2026Q2, өнөөдөр 2026-11-18) ═══")
    fake = FakeSession({(2026, 3): html_table(q3)})
    bs._SESSION = fake
    p, cands, head = bs.candidates_for((2026, 2), work / "r0", today=date(2026, 11, 18))
    check(p == (2026, 3) and len(cands) == 1, f"2026Q3 олдсон ({head})")
    check([c[:2] for c in fake.calls] == [(2026, 3)], f"зөвхөн дууссан улирлуудыг асуусан, шинээс эхэлж: {fake.calls}")
    check(all(c[2] == 1 for c in fake.calls), "report_type=1 (Монголбанкны стандарт)")
    fake2 = FakeSession({})
    bs._SESSION = fake2
    p2, c2, _ = bs.candidates_for((2026, 3), work / "r1", today=date(2027, 2, 20))
    check(p2 is None and [c[:2] for c in fake2.calls] == [(2026, 4), (2026, 3)], f"Q4 хараахан гараагүй үед: {fake2.calls}")

    print("\n═══ 3. Pipeline (Q3 workbook → Богд банк 2026/09/30) ═══")
    bs._SESSION = FakeSession({(2026, 3): html_table(q3)})
    bs._last_closed_quarter = lambda today=None: (2026, 3)
    cs.http_get = lambda url, timeout=60: Resp(js=[], headers={"X-WP-TotalPages": "1"})

    def no_ts(*a, **k):
        raise ConnectionError("offline")
    ts.http_get = no_ts
    bot.safe_get = no_ts
    wbk = work / "W.xlsx"
    shutil.copy(q3_wb, wbk)
    before = load_workbook(wbk)["Богд банк баланс"].max_column
    rc = qp.main(["--workbook", str(wbk), "--reports-dir", str(work / "r"), "--state", str(work / "s.json"),
                  "--analysis-dir", str(work / "an"), "--archive-dir", str(work / "arch")])
    check(rc == 0, f"exit {rc}")
    wf, wv = load_workbook(wbk), load_workbook(wbk, data_only=True)
    ws, wsv = wf["Богд банк баланс"], wv["Богд банк баланс"]
    c = ws.max_column
    check(c == before + 1 and ws.cell(4, c).value == "2026/09/30", "шинэ багана 2026/09/30")
    for lab in ("НИЙТ ХӨРӨНГИЙН ДҮН", "Харилцах", "Чанаргүй зээл", "ӨӨРИЙН ХӨРӨНГИЙН ДҮН", "Хугацаа хэтэрсэн зээл"):
        r = next(rr for rr in range(5, ws.max_row + 1) if str(ws.cell(rr, 1).value).strip() == lab)
        g, pv = wsv.cell(r, c).value, wsv.cell(r, c - 1).value
        check(isinstance(g, (int, float)) and abs(g / pv - 1.03) < 0.002, f"{lab}: {g and round(g):,} сая ₮ (мянгаас хөрвөсөн)")
    wi = wv["Богд банк орлого"]
    check(abs(wi.cell(13, wi.max_column).value / wi.cell(13, wi.max_column - 1).value - 1.5) < 0.002, "9 сарын ашиг")
    t = wf["9 банк ашиг"]
    row = next(r for r in range(6, 16) if t.cell(r, 1).value == "Богд банк")
    hdr = next(cc for cc in range(2, t.max_column + 1) if t.cell(5, cc).value == "2026/09")
    check(t.cell(row, hdr).value == f"='Богд банк орлого'!{bot.col_letter(wf['Богд банк орлого'].max_column)}13",
          f"trend: {t.cell(row, hdr).value}")
    check(str(wf["Тайлбар"]["B5"].value).startswith("Хамрах хугацаа: 2023/12/31"), "Тайлбар B5 хөндөгдөөгүй")
    check(qp.formula_errors(str(wbk)) <= qp.formula_errors(q3_wb), "томьёоны шинэ алдаагүй")
    m = json.loads((work / "an" / "2026Q3" / "metrics.json").read_text(encoding="utf-8"))["metrics"]["Богд банк"]
    check(m["date"] == "2026/09/30", f"шинжилгээнд Богд банк 2026/09/30 (ROE {m['roe']:.1%})")

    print("\n═══ 4. Дахин ажиллуулахад (ижил хариу) ═══")
    qp.main(["--workbook", str(wbk), "--reports-dir", str(work / "r"), "--state", str(work / "s.json"),
             "--analysis-dir", str(work / "an"), "--archive-dir", str(work / "arch")])
    check(load_workbook(wbk)["Богд банк баланс"].max_column == before + 1, "давхар багана нэмээгүй")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
