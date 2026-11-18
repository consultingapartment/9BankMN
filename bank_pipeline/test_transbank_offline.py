#!/usr/bin/env python3
"""
test_transbank_offline.py — transbank_source.py + pipeline-ийг сүлжээгүй шалгана.
    python3 test_transbank_offline.py <Q3_workbook.xlsx>
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402
import capitron_source as cs  # noqa: E402
import quarterly_pipeline as qp  # noqa: E402
import transbank_source as ts  # noqa: E402
from test_capitron_offline import Resp, make_pdf, make_xlsx  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


STD = ["Мөнгө ба түүнтэй адилтгах хөрөнгө", "Банк, СБ-д байршуулсан хөрөнгө", "Хөрөнгө оруулалт",
       "Зээл (цэвэр дүнгээр)", "Хэвийн зээл", "Чанаргүй зээл", "Зээлийн эрсдэлийн сан",
       "Дериватив санхүүгийн хөрөнгө", "Бусад санхүүгийн хөрөнгө", "Үндсэн хөрөнгө",
       "Борлуулах зориулалттай хөрөнгө", "Биет бус хөрөнгө", "НИЙТ ХӨРӨНГИЙН ДҮН", "Харилцах", "Хадгаламж",
       "Банк, СБ-аас татсан эх үүсвэр", "Бусад эх үүсвэр", "Бусад өр төлбөр", "НИЙТ ӨР ТӨЛБӨРИЙН ДҮН",
       "Хувь нийлүүлсэн хөрөнгө", "Нэмж төлөгдсөн капитал", "Хуримтлагдсан ашиг, алдагдал",
       "ӨӨРИЙН ХӨРӨНГИЙН ДҮН"]
INC = ["Хүүгийн орлого", "Хүүгийн зардал", "Хүүгийн цэвэр орлого", "Эрсдэлийн сангийн зардал",
       "Татварын өмнөх ашиг", "Татварын дараах ашиг"]


def rows_from(wb_vals, date, bm, im):
    wb, wi = wb_vals["Транс банк баланс"], wb_vals["Транс банк орлого"]
    cb, ci = qp.find_date_col(wb, 4, date), qp.find_date_col(wi, 4, date)
    out = []
    for ws, labs, c, m in ((wb, STD, cb, bm), (wi, INC, ci, im)):
        for lab in labs:
            r = qp.find_row(ws, 4, [lab])
            v = ws.cell(r, c).value if r else None
            if isinstance(v, (int, float)):
                out.append((lab, v * m))
    return out


def next_html(props, build_id="NEWBUILD123"):
    nd = {"props": {"pageProps": props}, "page": "/about-us/[...slug]", "buildId": build_id}
    return f'<html><body><div id="__next"></div><script id="__NEXT_DATA__" type="application/json">{json.dumps(nd, ensure_ascii=False)}</script></body></html>'


def main(q3_wb):
    work = Path(tempfile.mkdtemp(prefix="trans_test_"))
    ref = load_workbook(q3_wb, data_only=True)
    q3 = rows_from(ref, "2026/06/30", 1.05, 1.5)
    q2 = rows_from(ref, "2026/06/30", 1.0, 1.0)
    files = {
        f"{ts.SITE}/uploads/TB_FS_q3_8f2a1.pdf": make_pdf(q3, "ТРАНС БАНК — САНХҮҮГИЙН БАЙДЛЫН ТАЙЛАН", thousands=False),
        "https://cdn.transbank.mn/uploads/fs_2026_q2.xlsx": make_xlsx(q2),
        f"{ts.SITE}/files/2025-annual-audit.pdf": make_pdf(q2, "audit"),
    }
    props = {
        "page": {"title": "Улирлын санхүүгийн тайлан", "updatedAt": "2026-11-01T00:00:00.000Z"},
        "reports": [
            {"id": 31, "title": "2026 оны 3-р улирлын санхүүгийн тайлан", "createdAt": "2025-10-28T09:00:00.000Z",
             "file": {"data": {"attributes": {"url": "/uploads/TB_FS_q3_8f2a1.pdf", "ext": ".pdf",
                                               "mime": "application/pdf", "name": "TB_FS_q3_8f2a1.pdf"}}}},
            {"id": 30, "year": 2026, "quarter": 2, "title": "Санхүүгийн тайлан",
             "files": [{"url": "https://cdn.transbank.mn/uploads/fs_2026_q2.xlsx"}]},
        ],
        "body": '<p>Жилийн тайлан: <a href="/files/2025-annual-audit.pdf">2025 аудит</a></p>',
    }
    mode = {"html_has_props": True}

    def fake_get(url, timeout=60, extra=None):
        if url in ts.PAGES:
            return Resp(content=b"", js=None, headers={}) if False else _html(url)
        if "/_next/data/NEWBUILD123/mn/" in url:
            return Resp(js={"pageProps": props, "__N_SSG": True})
        if url in files:
            return Resp(content=files[url])
        return Resp(status=404)

    def _html(url):
        r = Resp(content=b"")
        r.text = next_html(props if mode["html_has_props"] else {})
        return r
    ts.http_get = fake_get
    cs.http_get = lambda url, timeout=60: Resp(js=[], headers={"X-WP-TotalPages": "1"})

    def no_net(*a, **k):
        raise ConnectionError("offline test")
    bot.safe_get = no_net

    print("═══ 1. pageProps ба файл таних ═══")
    p, bid = ts.fetch_page_props()
    check(p == props and bid == "NEWBUILD123", "HTML-ийн __NEXT_DATA__-аас pageProps + buildId")
    mode["html_has_props"] = False
    p2, bid2 = ts.fetch_page_props()
    check(p2 == props and bid2 == "NEWBUILD123", "HTML-д pageProps хоосон → шинэ buildId-аар _next/data")
    mode["html_has_props"] = True
    reps = {r["name"]: r["period"] for r in ts.list_reports(props)}
    check(reps.get("TB_FS_q3_8f2a1.pdf") == (2026, 3), "Strapi {url, ext, mime} + гарчиг '3-р улирал' → 2026Q3 (createdAt 2025-ийг тоохгүй)")
    check(reps.get("fs_2026_q2.xlsx") == (2026, 2), "year/quarter талбар → 2026Q2")
    check(reps.get("2025-annual-audit.pdf") == (2025, 4), "rich-text <a href> + 'annual audit' → 2025Q4")

    print("\n═══ 2. Pipeline (Q3 workbook → Транс банк 2026/09/30) ═══")
    wbk = work / "W.xlsx"
    shutil.copy(q3_wb, wbk)
    before = load_workbook(wbk)["Транс банк баланс"].max_column
    rc = qp.main(["--workbook", str(wbk), "--reports-dir", str(work / "r"), "--state", str(work / "s.json"),
                  "--analysis-dir", str(work / "an"), "--archive-dir", str(work / "arch")])
    check(rc == 0, f"exit {rc}")
    wf, wv = load_workbook(wbk), load_workbook(wbk, data_only=True)
    ws, wsv = wf["Транс банк баланс"], wv["Транс банк баланс"]
    c = ws.max_column
    check(c == before + 1 and ws.cell(4, c).value == "2026/09/30", "шинэ багана 2026/09/30")
    for lab in ("НИЙТ ХӨРӨНГИЙН ДҮН", "Хадгаламж", "Чанаргүй зээл", "ӨӨРИЙН ХӨРӨНГИЙН ДҮН", "Зээлийн эрсдэлийн сан"):
        r = qp.find_row(ws, 4, [lab]) or qp.find_row(ws, 4, ["  " + lab])
        r = r or next(rr for rr in range(5, ws.max_row + 1) if str(ws.cell(rr, 1).value).strip() == lab)
        g, pv = wsv.cell(r, c).value, wsv.cell(r, c - 1).value
        check(isinstance(g, (int, float)) and abs(g / pv - 1.05) < 0.002, f"{lab}: {g and round(g):,}")
    t = wf["Ашгийн зэрэгцүүлэлт"]
    row = next(r for r in range(6, 16) if t.cell(r, 1).value == "Транс банк")
    hdr = next(cc for cc in range(2, t.max_column + 1) if t.cell(5, cc).value == "2026/09")
    check(t.cell(row, hdr).value == f"='Транс банк орлого'!{bot.col_letter(wf['Транс банк орлого'].max_column)}13",
          f"trend: {t.cell(row, hdr).value}")
    check(str(wf["Тайлбар"]["B5"].value).startswith("Хамрах хугацаа: 2023/12/31"), "Тайлбар B5 хөндөгдөөгүй")
    check(qp.formula_errors(str(wbk)) <= qp.formula_errors(q3_wb), "томьёоны шинэ алдаагүй")
    m = json.loads((work / "an" / "2026Q3" / "metrics.json").read_text(encoding="utf-8"))["metrics"]["Транс банк"]
    check(m["date"] == "2026/09/30", f"шинжилгээнд Транс банк 2026/09/30 (NPL {m['npl_ratio']:.1%})")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
