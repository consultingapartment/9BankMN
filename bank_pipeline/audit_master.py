#!/usr/bin/env python3
"""audit_master.py — master Excel-ийн бүтэц/өгөгдлийн аудит (томьёоны алдаа, баланс тэнцэл,
огнооны дараалал, огцом үсрэлт, хоосон лавлагаа)."""
import re
import sys
from collections import Counter

from openpyxl import load_workbook

path = sys.argv[1]
wf = load_workbook(path)                    # томьёо
wv = load_workbook(path, data_only=True)    # кэшлэгдсэн утга
BANKS = ["Төрийн банк", "ХасБанк", "ХХБ", "Голомт", "Хаан банк", "Капитрон", "Ариг банк", "Богд банк", "Транс банк"]

print("=== 1. Томьёоны алдааны утга (#REF!, #DIV/0! г.м.) ===")
errs = Counter()
ex = {}
for ws in wv.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#") and c.value.rstrip("!?").upper() in (
                    "#REF", "#DIV/0", "#VALUE", "#NAME", "#N/A", "#NUM", "#NULL"):
                errs[ws.title] += 1
                ex.setdefault(ws.title, c.coordinate)
print(dict(errs) or "алдаагүй", ex or "")

print("\n=== 2. Томьёо хоосон нүд заасан эсэх (кэшлэгдсэн утга None/текст) ===")
n_total = 0
for ws in wf.worksheets:
    bad = []
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                n_total += 1
                v = wv[ws.title][c.coordinate].value
                if v is None:
                    bad.append(c.coordinate)
    if bad:
        print(f"  {ws.title}: {len(bad)} томьёо кэшлэгдсэн утгагүй, жишээ {bad[:4]}")
print("  нийт томьёо:", n_total)

print("\n=== 3. Баланс тэнцэл: Нийт хөрөнгө = Өр төлбөр + ӨХ (бүх банк, бүх багана) ===")
def find(ws, names):
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip().upper()
        if v in names:
            return r
for b in BANKS:
    ws = wv[f"{b} баланс"]
    hdr = next(r for r in range(1, 8) if str(ws.cell(r, 1).value).strip() == "Үзүүлэлт" and ws.cell(r, 2).value)
    ra = find(ws, {"НИЙТ ХӨРӨНГИЙН ДҮН"})
    rl = find(ws, {"НИЙТ ӨР ТӨЛБӨРИЙН ДҮН", "НИЙТ ӨР ТӨЛБӨР"})
    re_ = find(ws, {"ӨӨРИЙН ХӨРӨНГИЙН ДҮН", "НИЙТ ӨӨРИЙН ХӨРӨНГӨ"})
    bad = []
    for c in range(2, ws.max_column + 1):
        a, l, e = (ws.cell(r, c).value if r else None for r in (ra, rl, re_))
        if all(isinstance(x, (int, float)) for x in (a, l, e)):
            if abs(a - l - e) > max(1.0, 0.0005 * abs(a)):
                bad.append((ws.cell(hdr, c).value, round(a - l - e, 1)))
    print(f"  {b:12} {'OK' if not bad else 'ЗӨРҮҮ ' + str(bad)}")

print("\n=== 4. Огнооны толгой: давхардал, дараалал ===")
for b in BANKS:
    for kind in ("баланс", "орлого"):
        ws = wv[f"{b} {kind}"]
        hdr = next((r for r in range(1, 8) if str(ws.cell(r, 1).value).strip() == "Үзүүлэлт" and ws.cell(r, 2).value), None)
        heads = [str(ws.cell(hdr, c).value or "").strip().rstrip("*") for c in range(2, ws.max_column + 1)]
        heads = [h for h in heads if h]
        dup = [h for h, n in Counter(heads).items() if n > 1]
        ordered = heads == sorted(heads)
        if dup or not ordered:
            print(f"  {b} {kind}: давхардал={dup} хронологи={'тийм' if ordered else 'ҮГҮЙ'}  ({heads[0]}…{heads[-1]})")

print("\n=== 5. Утгын шалгалт ===")
import statistics
def cols_of(ws):
    hdr = next((r for r in range(1, 8) if str(ws.cell(r, 1).value).strip() == "Үзүүлэлт" and ws.cell(r, 2).value), None)
    cs = sorted(((str(ws.cell(hdr, c).value or "").strip().rstrip("*"), c) for c in range(2, ws.max_column + 1) if ws.cell(hdr, c).value))
    return hdr, cs
print("  (а) Баланс: хөрш хугацаанаас >60% ба >5 мянга зөрсөн")
for b in BANKS:
    ws = wv[f"{b} баланс"]; hdr, cs = cols_of(ws)
    for r in range(hdr + 1, ws.max_row + 1):
        lab = str(ws.cell(r, 1).value or "").strip()
        if lab.lower() not in ("нийт хөрөнгийн дүн", "зээл (цэвэр дүнгээр)", "хадгаламж", "харилцах", "хэвийн зээл", "чанаргүй зээл", "чанаргүй зээл (npl)"):
            continue
        vals = [(h, ws.cell(r, c).value) for h, c in cs if isinstance(ws.cell(r, c).value, (int, float))]
        for (h1, v1), (h2, v2) in zip(vals, vals[1:]):
            if abs(v1) > 5000 and abs(v2) > 5000 and abs(v2 - v1) / abs(v1) > 0.6:
                print(f"     {b} {lab}: {h1}={v1:,.0f} → {h2}={v2:,.0f}")
print("  (б) Хуримтлалын орлого (ХасБанкаас бусад): оны дотор буурсан мөр")
for b in BANKS:
    if b == "ХасБанк":
        continue
    ws = wv[f"{b} орлого"]; hdr, cs = cols_of(ws)
    for r in range(hdr + 1, ws.max_row + 1):
        lab = str(ws.cell(r, 1).value or "").strip()
        if lab.lower() not in ("хүүгийн орлого", "хүүгийн зардал", "хүүгийн бус зардал"):
            continue
        vals = [(h, ws.cell(r, c).value) for h, c in cs if isinstance(ws.cell(r, c).value, (int, float))]
        for (h1, v1), (h2, v2) in zip(vals, vals[1:]):
            if h1[:4] == h2[:4] and v1 > 0 and v2 < v1 * 0.995:
                print(f"     {b} {lab}: {h1}={v1:,.0f} → {h2}={v2:,.0f} (хуримтлал буурсан)")
print("  (в) ХасБанк орлого (улирлын дан дүн): сүүлийн 8 улирлын медианаас >2.5x эсвэл <0.4x")
ws = wv["ХасБанк орлого"]; hdr, cs = cols_of(ws)
cs = [(h, c) for h, c in cs if not h.endswith("12/31") or h[:4] >= "2025"]  # 2023/12, 2024/12 нь жилийн дүн байж болзошгүй
for r in range(hdr + 1, ws.max_row + 1):
    lab = str(ws.cell(r, 1).value or "").strip()
    vals = [(h, ws.cell(r, c).value) for h, c in cs if isinstance(ws.cell(r, c).value, (int, float)) and h >= "2025"]
    if len(vals) < 4:
        continue
    med = statistics.median(abs(v) for _, v in vals)
    for h, v in vals:
        if med > 1000 and (abs(v) > 2.5 * med or abs(v) < 0.4 * med):
            print(f"     ХасБанк орлого {lab}: {h}={v:,.0f} (медиан {med:,.0f})")

print("\n=== 6. Гарчиг/тэмдэглэл дэх огноо, 'FY' шошго ===")
d = wf["Дашборд"]
for r in range(1, 12):
    v = d.cell(r, 1).value
    if v and ("FY" in str(v) or "2026" in str(v)):
        print(f"  Дашборд A{r}: {v}")
