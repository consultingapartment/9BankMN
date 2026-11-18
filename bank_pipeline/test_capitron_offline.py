#!/usr/bin/env python3
"""
test_capitron_offline.py — capitron_source.py + pipeline-ийг сүлжээгүй шалгана.
WordPress API-ийн хариу, PDF/xlsx файлыг дуурайлгаж (HTTP mock), Q3 workbook дээр
Капитроны 2026/09/30 баганыг нэмүүлнэ.

    python3 test_capitron_offline.py <Q3_workbook.xlsx>
"""
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402
import capitron_source as cs  # noqa: E402
import quarterly_pipeline as qp  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


# Монголбанкны стандарт маягтын шошгууд (Капитрон sheet-ийн нэрээс ӨӨР) -> sheet-ийн мөр
STD_BAL = [
    ("Мөнгө ба түүнтэй адилтгах хөрөнгө", "Мөнгө ба түүнтэй адилтгах хөрөнгө", 1.0),
    ("Хөрөнгө оруулалт", "Хөрөнгө оруулалт (санхүүгийн)", 1.0),
    ("Зээл (цэвэр дүнгээр)", "Зээл (цэвэр дүнгээр)", 1.0),
    ("Хэвийн зээл", "Хэвийн зээл", 1.0),
    ("Хугацаа хэтэрсэн зээл", "Хугацаа хэтэрсэн зээл", 1.0),
    ("Хэвийн бус зээл", "Чанаргүй зээл (NPL)", 0.3),
    ("Эргэлзээтэй зээл", "Чанаргүй зээл (NPL)", 0.3),
    ("Муу зээл", "Чанаргүй зээл (NPL)", 0.4),
    ("Зээлийн эрсдэлийн сан", "Зээлийн эрсдэлийн сан", 1.0),
    ("Дериватив санхүүгийн хөрөнгө", "Дериватив санхүүгийн хөрөнгө", 1.0),
    ("Бусад санхүүгийн хөрөнгө", "Бусад хөрөнгө", 0.6),
    ("Бусад санхүүгийн бус хөрөнгө", "Бусад хөрөнгө", 0.4),
    ("Үндсэн хөрөнгө", "Үндсэн хөрөнгө", 1.0),
    ("Борлуулах зориулалттай хөрөнгө", "Борлуулах зориулалттай хөрөнгө", 1.0),
    ("Биет бус хөрөнгө", "Биет бус хөрөнгө", 1.0),
    ("НИЙТ ХӨРӨНГИЙН ДҮН", "НИЙТ ХӨРӨНГИЙН ДҮН", 1.0),
    ("Харилцах", "Харилцагчдын харилцах, хадгаламж", 0.35),
    ("Хадгаламж", "Харилцагчдын харилцах, хадгаламж", 0.65),
    ("Банк, санхүүгийн байгууллагаас татсан эх үүсвэр", "Банк, СБ-аас татсан эх үүсвэр", 1.0),
    ("Бусад өр төлбөр", "Бусад өр төлбөр (репо, бусад)", 1.0),
    ("НИЙТ ӨР ТӨЛБӨРИЙН ДҮН", "НИЙТ ӨР ТӨЛБӨРИЙН ДҮН", 1.0),
    ("Хувь нийлүүлсэн хөрөнгө", "Хувь нийлүүлсэн хөрөнгө", 1.0),
    ("Хуримтлагдсан ашиг", "Хуримтлагдсан ашиг, алдагдал", 1.0),
    ("ӨӨРИЙН ХӨРӨНГИЙН ДҮН", "ӨӨРИЙН ХӨРӨНГИЙН ДҮН", 1.0),
]
STD_INC = ["Хүүгийн орлого", "Хүүгийн зардал", "Хүүгийн цэвэр орлого", "Эрсдэлийн сангийн зардал",
           "Татварын өмнөх ашиг", "Орлогын татварын зардал", "Татварын дараах ашиг"]


def source_values(wb_vals, date, bal_mult, inc_mult):
    wb, wi = wb_vals["Капитрон баланс"], wb_vals["Капитрон орлого"]
    cb, ci = qp.find_date_col(wb, 4, date), qp.find_date_col(wi, 4, date)

    def v(ws, c, label):
        r = qp.find_row(ws, 4, [label])
        x = ws.cell(r, c).value if r else None
        return x if isinstance(x, (int, float)) else None
    rows = []
    for std, sheet_lab, share in STD_BAL:
        x = v(wb, cb, sheet_lab)
        if x is not None:
            rows.append((std, x * share * bal_mult))
    for lab in STD_INC:
        x = v(wi, ci, lab)
        if x is not None:
            rows.append((lab, x * inc_mult))
    return rows


def make_pdf(rows, title, thousands=True) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas
    pdfmetrics.registerFont(TTFont("DV", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setFont("DV", 10)
    y = 800
    c.drawString(40, y, title)
    y -= 16
    c.drawString(40, y, "(мянган төгрөгөөр)" if thousands else "(сая төгрөгөөр)")
    y -= 20
    c.drawString(40, y, "№  Үзүүлэлт")
    c.drawString(380, y, "Өмнөх үе")
    c.drawString(480, y, "Тайлант үе")
    y -= 16
    m = 1000 if thousands else 1
    for i, (lab, val) in enumerate(rows, 1):
        if y < 60:
            c.showPage(); c.setFont("DV", 10); y = 800
        c.drawString(40, y, f"{i}. {lab}")
        prev, cur = val * 0.96 * m, val * m

        def fmt(x):
            s = f"{abs(x):,.1f}"
            return f"({s})" if x < 0 else s
        c.drawRightString(460, y, fmt(prev))
        c.drawRightString(560, y, fmt(cur))
        y -= 14
    c.save()
    return buf.getvalue()


def make_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(["Санхүүгийн тайлан 2026.06.30"])
    for lab, v in rows:
        ws.append([lab, v * 0.97, v])
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


class Resp:
    def __init__(self, content=b"", js=None, headers=None, status=200):
        self.content, self._js, self.headers, self.status_code = content, js, headers or {}, status

    def json(self):
        return self._js

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")


def main(q3_wb):
    work = Path(tempfile.mkdtemp(prefix="cap_test_"))
    ref = load_workbook(q3_wb, data_only=True)
    q3_rows = source_values(ref, "2026/06/30", 1.04, 1.5)       # Q3 = H1-ээс зохиомол (YTD ×1.5)
    q2_rows = source_values(ref, "2026/06/30", 1.0, 1.0)
    U = "https://www.capitronbank.mn/wp-content/uploads"
    files = {
        f"{U}/2026/10/Capitron_Q3_2026.pdf": make_pdf(q3_rows, "КАПИТРОН БАНК — САНХҮҮГИЙН ТАЙЛАН 2026.09.30"),
        f"{U}/2026/07/report-2026-06.xlsx": make_xlsx(q2_rows),
        f"{U}/2026/04/audit2025.pdf": make_pdf(q2_rows, "Аудитлагдсан тайлан 2025"),
    }
    posts43 = [
        {"id": 901, "date": "2026-10-30T10:00:00", "link": "x",
         "title": {"rendered": "2026 оны III улирлын санхүүгийн тайлан"},
         "content": {"rendered": f'<p>Тайлан: <a href="{U}/2026/10/Capitron_Q3_2026.pdf">татах</a></p>'}},
        {"id": 902, "date": "2026-07-30T10:00:00", "link": "x",
         "title": {"rendered": "2026 оны 2-р улирлын тайлан"}, "content": {"rendered": "<p></p>"},
         "_embedded": {"wp:featuredmedia": [{"source_url": f"{U}/2026/07/report-2026-06.xlsx"}]}},
        {"id": 903, "date": "2026-01-10T10:00:00", "link": "x",
         "title": {"rendered": "Шинэ салбар нээгдлээ"}, "content": {"rendered": "<p>мэдээ</p>"}},
    ]
    posts44 = [
        {"id": 904, "date": "2026-04-20T10:00:00", "link": "x",
         "title": {"rendered": "2025 оны аудитлагдсан санхүүгийн тайлан"},
         "content": {"rendered": ""}, "acf": {"file": {"url": f"{U}/2026/04/audit2025.pdf"}}},
    ]

    def fake_get(url, timeout=60):
        if "/posts?" in url:
            cat = 43 if "categories=43" in url else 44
            return Resp(js=posts43 if cat == 43 else posts44, headers={"X-WP-TotalPages": "1"})
        if "/media?" in url:
            return Resp(js=[])
        if url in files:
            return Resp(content=files[url])
        return Resp(status=404)
    cs.http_get = fake_get

    def no_net(*a, **k):
        raise ConnectionError("offline test")
    bot.safe_get = no_net

    print("═══ 1. Жагсаалт ба улирал таних ═══")
    reps = cs.list_reports()
    per = {r["title"]: r["period"] for r in reps}
    check(per["2026 оны III улирлын санхүүгийн тайлан"] == (2026, 3), "'III улирлын' → 2026Q3")
    check(per["2026 оны 2-р улирлын тайлан"] == (2026, 2), "'2-р улирлын' → 2026Q2 (_embedded xlsx)")
    check(per["2025 оны аудитлагдсан санхүүгийн тайлан"] == (2025, 4), "'аудитлагдсан' → 2025Q4 (acf талбарын файл)")
    check(per["Шинэ салбар нээгдлээ"] is None, "тайлан биш нийтлэл → улиралгүй")
    for t, exp in [("2026.09.30-ны байдлаар", (2026, 3)), ("2026 оны 6 сарын 30", (2026, 2)),
                   ("Q4 2025", (2025, 4)), ("2026 оны хагас жилийн", (2026, 2)), ("2026 оны 9 сарын тайлан", (2026, 3))]:
        check(cs.detect_period(t) == exp, f"detect_period('{t}') = {exp}")

    print("\n═══ 2. PDF задлалт (мянган ₮, хаалттай сөрөг тоо) ═══")
    p = work / "q3.pdf"
    p.write_bytes(files[f"{U}/2026/10/Capitron_Q3_2026.pdf"])
    raw = cs.adapt_to_sheet(cs.parse_pdf(p))
    exp = dict(q3_rows)
    check(abs(raw["НИЙТ ХӨРӨНГИЙН ДҮН"] - exp["НИЙТ ХӨРӨНГИЙН ДҮН"] * 1000) < 1, "нийт хөрөнгө = баруун багана")
    check(raw["Зээлийн эрсдэлийн сан"] < 0, "(…) → сөрөг")
    check("Харилцагчдын харилцах, хадгаламж" in raw, "Харилцах + Хадгаламж → Капитрон sheet-ийн нэгдсэн мөр")
    check("Чанаргүй зээл" in raw, "Хэвийн бус + Эргэлзээтэй + Муу → Чанаргүй зээл")

    print("\n═══ 3. Pipeline (Q3 workbook → Капитрон 2026/09/30) ═══")
    wbk = work / "W.xlsx"
    shutil.copy(q3_wb, wbk)
    before = load_workbook(wbk)["Капитрон баланс"].max_column
    rc = qp.main(["--workbook", str(wbk), "--reports-dir", str(work / "reports"), "--state", str(work / "s.json"),
                  "--analysis-dir", str(work / "an"), "--archive-dir", str(work / "arch")])
    check(rc == 0, f"exit {rc}")
    wf, wv = load_workbook(wbk), load_workbook(wbk, data_only=True)
    ws, wsv = wf["Капитрон баланс"], wv["Капитрон баланс"]
    check(ws.max_column == before + 1 and ws.cell(4, ws.max_column).value == "2026/09/30", "шинэ багана 2026/09/30")
    c = ws.max_column
    for lab in ("НИЙТ ХӨРӨНГИЙН ДҮН", "Харилцагчдын харилцах, хадгаламж", "Чанаргүй зээл (NPL)", "Бусад хөрөнгө",
                "Хөрөнгө оруулалт (санхүүгийн)", "ӨӨРИЙН ХӨРӨНГИЙН ДҮН"):
        r = qp.find_row(ws, 4, [lab])
        got = wsv.cell(r, c).value
        prev = wsv.cell(r, c - 1).value
        ok = isinstance(got, (int, float)) and isinstance(prev, (int, float)) and abs(got / prev - 1.04) < 0.002
        check(ok, f"{lab}: {got and round(got):,} (= H1 × 1.04, сая ₮ болж хөрвөсөн)")
    wi = wv["Капитрон орлого"]
    r = qp.find_row(wi, 4, ["Татварын дараах ашиг"])
    check(abs(wi.cell(r, wi.max_column).value / wi.cell(r, wi.max_column - 1).value - 1.5) < 0.002, "9 сарын ашиг (YTD)")
    t = wf["9 банк зэрэгцүүлэлт"]
    row = next(r for r in range(6, 16) if t.cell(r, 1).value == "Капитрон")
    hdr = next(c for c in range(2, t.max_column + 1) if t.cell(5, c).value == "2026/09")
    check(t.cell(row, hdr).value == f"='Капитрон баланс'!{bot.col_letter(c)}22", f"trend: 2026/09 баганад ({t.cell(row, hdr).value}), шинэ багана нэмээгүй")
    check(all(t.cell(h, hdr).value == "2026/09" for h in (16, 27)), "доод блокийн 'Банк / Улирал' толгойнууд бөглөгдсөн")
    cam = [x.value for row_ in wf["CAMELS дүгнэлт"].iter_rows() for x in row_
           if isinstance(x.value, str) and "'Капитрон орлого'!" in x.value]
    check(cam and all(f"!{bot.col_letter(wf['Капитрон орлого'].max_column)}" in f for f in cam), "CAMELS → шинэ багана руу чиглэсэн")
    check(any("*(4/3)" in f for f in cam), "Капитрон жилчлэл ×4/3")
    check(str(wf["Тайлбар"]["B5"].value).startswith("Хамрах хугацаа: 2023/12/31"), "Тайлбар B5: хамрах хугацааны эхлэл хөндөгдөөгүй")
    check(qp.formula_errors(str(wbk)) <= qp.formula_errors(q3_wb), "томьёоны шинэ алдаагүй")
    m = json.loads((work / "an" / "2026Q3" / "metrics.json").read_text(encoding="utf-8"))["metrics"]["Капитрон"]
    check(m["date"] == "2026/09/30", f"шинжилгээнд Капитрон 2026/09/30 (ROE {m['roe']:.1%})")

    print("\n═══ 4. Дахин ажиллуулахад ═══")
    qp.main(["--workbook", str(wbk), "--reports-dir", str(work / "reports"), "--state", str(work / "s.json"),
             "--analysis-dir", str(work / "an"), "--archive-dir", str(work / "arch")])
    check(load_workbook(wbk)["Капитрон баланс"].max_column == before + 1, "давхар багана нэмээгүй")

    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
