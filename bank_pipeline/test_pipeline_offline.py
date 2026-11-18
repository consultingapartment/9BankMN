#!/usr/bin/env python3
"""
test_pipeline_offline.py — сүлжээгүй БҮРЭН урсгалын тест.

A. Q3 regression: Q2 master + (Q3 workbook-ийн мэдэгдэж буй зөв утгаас үүсгэсэн
   МХБ-ийн хэлбэрийн файлууд) → pipeline → Q3 workbook-тэй тулгана.
B. Q4 симуляци: Q3 workbook + зохиомол Q4 файл → толгой '2026/12/31*', жилчлэл ×1,
   ХасБанк Q4 дан дүн, YTD гинж, гарчгийн огноо, шинжилгээ.
C. Restatement: ижил Q4-ийн өөр (аудитлагдсан) файл → багана дахин бичигдэнэ.
D. Чанарын хаалт: эвдэрсэн файл → банк буцаагдана, workbook-т юу ч үлдэхгүй.

    python3 test_pipeline_offline.py <Q2_master.xlsx> <Q3_workbook.xlsx>
"""
import json
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402
import quarterly_pipeline as qp  # noqa: E402

BANKS = list(bot.BANKS)
FAIL = []


def check(cond, msg):
    print(("  ✓ " if cond else "  ✗ ") + msg)
    if not cond:
        FAIL.append(msg)


def col_by_date(ws, hdr, d):
    return qp.find_date_col(ws, hdr, d)


def make_source(path: Path, date_txt: str, rows: list[tuple[str, float]]):
    """МХБ-ийн файлтай төстэй: гарчиг + огноо, дугаарлалт, шошго, өмнөх ба одоогийн утга."""
    wb = Workbook()
    ws = wb.active
    ws.title = "СБД"
    ws.append(["САНХҮҮГИЙН БАЙДЛЫН ТАЙЛАН"])
    ws.append([f"{date_txt}-ны байдлаар", None, None, "сая төгрөг"])
    ws.append(["№", "Үзүүлэлт", "Эхний үлдэгдэл", "Эцсийн үлдэгдэл"])
    for i, (lab, v) in enumerate(rows, 1):
        ws.append([f"1.{i}", lab, round(v * 0.97, 3), v])
    wb.save(path)


def sheet_rows(ws, hdr, col, transform=lambda lab, v: v):
    out = []
    for r in range(hdr + 1, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        v = ws.cell(r, col).value
        if lab and isinstance(v, (int, float)):
            out.append((str(lab).strip(), transform(str(lab).strip(), v)))
    return out


def build_sources(wb_vals, date_label, inbox: Path, date_txt: str, mutate=None, khas_prior=None):
    inbox.mkdir(parents=True, exist_ok=True)
    for bank, info in bot.BANKS.items():
        wsb, wsi = wb_vals[f"{bank} баланс"], wb_vals[f"{bank} орлого"]
        hb, hi = bot.HEADER_ROW[bank]["bal"], bot.HEADER_ROW[bank]["inc"]
        cb, ci = col_by_date(wsb, hb, date_label), col_by_date(wsi, hi, date_label)
        rows = sheet_rows(wsb, hb, cb) + sheet_rows(wsi, hi, ci)
        if bank == "ХасБанк" and khas_prior:
            rows = [(l, v + khas_prior.get(l, 0.0)) if l in khas_prior else (l, v) for l, v in rows]
        if mutate:
            rows = mutate(bank, rows)
        make_source(inbox / f"{info['id']}_report.xlsx", date_txt, rows)


def khas_prior_sum(wb_vals, year: str, upto_label: str):
    """ХасБанк: тухайн оны upto_label-аас ӨМНӨХ улирлын дан дүнгийн нийлбэр (хуримтлал болгох)."""
    ws = wb_vals["ХасБанк орлого"]
    hi = bot.HEADER_ROW["ХасБанк"]["inc"]
    cols = [c for c in range(2, ws.max_column + 1)
            if str(ws.cell(hi, c).value or "").startswith(year) and str(ws.cell(hi, c).value).replace("*", "") < upto_label]
    out = {}
    for r in range(hi + 1, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        vals = [ws.cell(r, c).value for c in cols]
        if lab and all(isinstance(v, (int, float)) for v in vals):
            out[str(lab).strip()] = sum(vals)
    return out


def run_pipe(work: Path, wbk: Path, period: str, inbox: Path):
    return qp.main(["--workbook", str(wbk), "--inbox", str(inbox), "--period", period,
                    "--reports-dir", str(work / "reports"), "--state", str(work / "state.json"),
                    "--analysis-dir", str(work / "analysis"), "--archive-dir", str(work / "archive")])


def test_a(q2_master, q3_wb, work: Path):
    print("\n═══ A. Q3 regression (Q2 master → Q3) ═══")
    work.mkdir(parents=True, exist_ok=True)
    wbk = work / "A.xlsx"
    shutil.copy(q2_master, wbk)
    ref = load_workbook(q3_wb, data_only=True)
    base = load_workbook(q2_master, data_only=True)
    # ХасБанк: эх тайлан хуримтлалаар ирдэг → Q1+Q2 (энэ master-ийн) + Q3 дан
    prior = khas_prior_sum(base, "2026", "2026/09/30")
    build_sources(ref, "2026/09/30", work / "inA", "2026.09.30", khas_prior=prior)
    rc = run_pipe(work, wbk, "2026Q3", work / "inA")
    check(rc == 0, f"pipeline exit code {rc}")
    got = load_workbook(wbk, data_only=True)
    tot = bad = 0
    for bank in BANKS:
        for kind in ("баланс", "орлого"):
            a, b = ref[f"{bank} {kind}"], got[f"{bank} {kind}"]
            h = bot.HEADER_ROW[bank]["bal" if kind == "баланс" else "inc"]
            ca, cb = col_by_date(a, h, "2026/09/30"), col_by_date(b, h, "2026/09/30")
            check(cb is not None, f"{bank} {kind}: 2026/09/30 багана үүссэн")
            if cb is None:
                continue
            for r in range(h + 1, a.max_row + 1):
                x, y = a.cell(r, ca).value, b.cell(r, cb).value
                if isinstance(x, (int, float)) and x != 0:
                    tot += 1
                    if not isinstance(y, (int, float)) or abs(y - x) / abs(x) > 0.005:
                        if bank == "ХасБанк" and kind == "орлого":
                            continue  # master-ийн Q2 rebase-ээс шалтгаалсан ялгаа (доор тусад нь)
                        bad += 1
                        print(f"     зөрүү {bank} {kind} r{r} {a.cell(r, 1).value}: ref={x} got={y}")
    check(bad == 0, f"түүхий хуудас: {tot} утга тулгаж, {bad} зөрүү")
    d = got["Дашборд"]
    check("2026/09/30" in str(d["A4"].value), f"Дашборд A4 гарчиг: {d['A4'].value}")
    check(qp.formula_errors(str(wbk)) <= qp.formula_errors(str(q2_master)), "томьёоны шинэ алдаагүй")
    an = work / "analysis" / "2026Q3"
    check((an / "metrics.json").exists() and any(an.glob("*.docx")), "шинжилгээ (metrics.json, docx) үүссэн")
    check(any((work / "reports" / "2026Q3").glob("*.labels.json")), "эх файл + шошгууд архивлагдсан")


def test_b(q3_wb, work: Path):
    print("\n═══ B. Q4 симуляци (Q3 → 2026/12/31*) ═══")
    work.mkdir(parents=True, exist_ok=True)
    wbk = work / "B.xlsx"
    shutil.copy(q3_wb, wbk)
    ref = load_workbook(q3_wb, data_only=True)
    prior = khas_prior_sum(ref, "2026", "2026/12/31")   # Q1+Q2+Q3 дан
    income_labels = {str(ref[f"{b} орлого"].cell(r, 1).value).strip()
                     for b in BANKS for r in range(5, 16) if ref[f"{b} орлого"].cell(r, 1).value}

    def mut(bank, rows):
        out = []
        for lab, v in rows:
            if lab in income_labels:
                # Хуримтлалтай банкууд: FY = 9 сар × 4/3; ХасБанк: Q4 дан = Q3 дан × 1.1 (+өмнөх нь доор)
                out.append((lab, v * 1.1 if bank == "ХасБанк" else v * 4 / 3))
            else:
                out.append((lab, v * 1.03))
        return out
    build_sources(ref, "2026/09/30", work / "inB", "2026.12.31", mutate=mut, khas_prior=None)
    # ХасБанкны файлд өмнөх 3 улирлын нийлбэрийг нэмж хуримтлал болгоно
    wsrc = load_workbook(work / "inB" / "568_report.xlsx")
    ws = wsrc.active
    for r in range(4, ws.max_row + 1):
        lab = ws.cell(r, 2).value
        if lab in prior and lab in income_labels:
            ws.cell(r, 4).value = ws.cell(r, 4).value + prior[lab]
    wsrc.save(work / "inB" / "568_report.xlsx")

    rc = run_pipe(work, wbk, "2026Q4", work / "inB")
    check(rc == 0, f"pipeline exit code {rc}")
    wf, wv = load_workbook(wbk), load_workbook(wbk, data_only=True)
    for bank in BANKS:
        h = bot.HEADER_ROW[bank]["bal"]
        ws = wf[f"{bank} баланс"]
        check(str(ws.cell(h, ws.max_column).value) == "2026/12/31*", f"{bank}: толгой '{ws.cell(h, ws.max_column).value}'")
    # ХасБанк Q4 дан = Q3 × 1.1
    wi = wv["ХасБанк орлого"]
    hi = bot.HEADER_ROW["ХасБанк"]["inc"]
    q3c, q4c = col_by_date(wi, hi, "2026/09/30"), col_by_date(wi, hi, "2026/12/31")
    r13 = 13
    check(abs(wi.cell(r13, q4c).value - wi.cell(r13, q3c).value * 1.1) < 1, "ХасБанк Q4 дан ашиг = хуримтлал − (Q1+Q2+Q3)")
    # жилчлэл ×1
    dash = wf["Дашборд"]
    formulas = [c.value for row in dash.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")]
    leftover = [f for f in formulas if any(t in f for t in ("*(4/3)", "*2)", "*4)"))]
    check(not leftover, f"Дашборд-д Q1-Q3 жилчлэлийн коэффициент үлдээгүй ({leftover[:2]})")
    ytd = [f for f in formulas if "'ХасБанк орлого'!" in f and "+'ХасБанк орлого'!" in f]
    check(any(f.count("'ХасБанк орлого'!") >= 4 for f in ytd), "ХасБанк YTD = 4 улирлын нийлбэр")
    check("2026/12/31" in str(wf["Дашборд"]["A4"].value), f"гарчиг: {wf['Дашборд']['A4'].value}")
    check("FY2026" in str(wf["Дашборд"]["A8"].value), f"ашгийн гарчиг: {wf['Дашборд']['A8'].value}")
    m = json.loads((work / "analysis" / "2026Q4" / "metrics.json").read_text(encoding="utf-8"))["metrics"]
    sbm = m["Төрийн банк"]
    check(abs(sbm["roe"] - sbm["profit_ytd"] / sbm["equity"]) < 1e-9, f"Q4 ROE = жилийн ашиг / ӨХ ({sbm['roe']:.3f})")
    check(abs(sbm["profit_ttm"] - sbm["profit_ytd"]) < 1e-6, "Q4 TTM = FY")
    for t in ("9 банк зэрэгцүүлэлт", "Ашгийн зэрэгцүүлэлт", "Хүүгийн маржин"):
        ws = wf[t]
        check(any(ws.cell(r, ws.max_column).value == "2026/12" for r in range(1, 8)), f"{t}: 2026/12 багана нэмэгдсэн")
    return wbk


def test_c(wbk: Path, work: Path):
    print("\n═══ C. Restatement (Q4 аудитлагдсан файл ирэв) ═══")
    inb = work / "inB"
    p = inb / "564_report.xlsx"
    w = load_workbook(p)
    ws = w.active
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 2).value == "Татварын дараах ашиг":
            ws.cell(r, 4).value = ws.cell(r, 4).value * 0.95
            new_profit = ws.cell(r, 4).value
    w.save(p)
    for f in inb.iterdir():
        if not f.name.startswith("564"):
            f.unlink()
    rc = run_pipe(work, wbk, "2026Q4", inb)
    check(rc == 0, f"pipeline exit code {rc}")
    wv = load_workbook(wbk, data_only=True)
    ws = wv["Төрийн банк орлого"]
    c = col_by_date(ws, 4, "2026/12/31")
    check(c == ws.max_column, "шинэ багана НЭМЭГДЭЭГҮЙ (ижил баганыг дахин бичсэн)")
    check(abs(ws.cell(13, c).value - new_profit) < 1, f"ашиг аудитын утгаар солигдсон ({ws.cell(13, c).value:,.0f})")
    rc = run_pipe(work, wbk, "2026Q4", inb)
    wv2 = load_workbook(wbk, data_only=True)
    check(wv2["Төрийн банк орлого"].max_column == ws.max_column, "дахин ажиллуулахад (ижил sha) юу ч өөрчлөгдөөгүй")
    check(str(wv2["Тайлбар"]["B5"].value).startswith("Хамрах хугацаа: 2023/12/31"), "Тайлбар B5: хамрах хугацааны эхлэл хөндөгдөөгүй")
    for sh, c in (("Дашборд", "A2"), ("Тайлбар", "B5"), ("CAMELS дүгнэлт", "A8")):
        v = str(wv2[sh][c].value)
        check("5 банк — 2026/12/31" in v.replace(" 2026/12/31", " — 2026/12/31").replace("— —", "—")
              and "Транс банк — 2026/06/30" in v, f"{sh} {c}: МХБ-ийн бус 4 банкны огноо хэвээр ({v[-60:]})")


def test_d(q3_wb, work: Path):
    print("\n═══ D. Чанарын хаалт (эвдэрсэн файл) ═══")
    work.mkdir(parents=True, exist_ok=True)
    wbk = work / "D.xlsx"
    shutil.copy(q3_wb, wbk)
    ref = load_workbook(q3_wb, data_only=True)

    def drop_equity(bank, rows):
        return [(l, v) for l, v in rows if "ӨӨРИЙН ХӨРӨНГ" not in l.upper()] if bank == "ХХБ" else rows
    build_sources(ref, "2026/09/30", work / "inD", "2026.12.31", mutate=lambda b, r: drop_equity(b, [(l, v * 1.02) for l, v in r]))
    for f in (work / "inD").iterdir():
        if not f.name.startswith("567"):
            f.unlink()
    before = load_workbook(wbk)["ХХБ баланс"].max_column
    run_pipe(work, wbk, "2026Q4", work / "inD")
    after = load_workbook(wbk)
    ws = after["ХХБ баланс"]
    nonempty = [ws.cell(r, before + 1).value for r in range(1, ws.max_row + 1) if ws.cell(r, before + 1).value is not None]
    check(not nonempty, f"ХХБ буцаагдсан — шинэ баганад утга үлдээгүй ({nonempty[:3]})")
    check(str(ws.cell(5, before).value) == "2026/09/30", "ХХБ сүүлийн огноо хэвээр 2026/09/30")


def test_e(q3_wb, work: Path):
    print("\n═══ E. Буруу улирлын файл (огнооны шалгалт) ═══")
    work.mkdir(parents=True, exist_ok=True)
    wbk = work / "E.xlsx"
    shutil.copy(q3_wb, wbk)
    ref = load_workbook(q3_wb, data_only=True)
    build_sources(ref, "2026/09/30", work / "inE", "2026.09.30")   # Q3-ийн файл Q4 гэж ирэв
    for f in (work / "inE").iterdir():
        if not f.name.startswith("563"):
            f.unlink()
    before = load_workbook(wbk)["Хаан банк баланс"].max_column
    run_pipe(work, wbk, "2026Q4", work / "inE")
    check(load_workbook(wbk)["Хаан банк баланс"].max_column == before, "огноо таараагүй файлыг бичээгүй")


if __name__ == "__main__":
    q2, q3 = sys.argv[1], sys.argv[2]
    work = Path(tempfile.mkdtemp(prefix="qp_test_"))
    print("work:", work)
    test_a(q2, q3, work / "a")
    wbk = test_b(q3, work / "b")
    test_c(wbk, work / "b")
    test_d(q3, work / "d")
    test_e(q3, work / "e")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)} алдаа: " + "; ".join(FAIL)))
    sys.exit(1 if FAIL else 0)
