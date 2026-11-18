#!/usr/bin/env python3
"""
capitron_source.py — Капитрон банкны улирлын тайланг capitronbank.mn-ийн WordPress
REST API-аас олж, татаж, (шошго -> утга) болгон задална.

Эх сурвалж (сайтын "Бидний тухай → Тайлан" хэсэг өөрөө ингэж дууддаг):
    https://www.capitronbank.mn/admin//wp-json/wp/v2/posts?_embed=true&categories=43&per_page=100
    https://www.capitronbank.mn/admin//wp-json/wp/v2/posts?_embed=true&categories=44&per_page=100

Нийтлэл (post) бүрээс:
    • гарчиг/агуулга/файлын нэрээс тайлант УЛИРЛЫГ тодорхойлно
      ("2026 оны II улирал", "2-р улирал", "2026.06.30", "6 сарын 30", "хагас жил",
       "жилийн эцсийн/аудитлагдсан" → Q4 гэх мэт)
    • .pdf/.xlsx/.xls хавсралтыг ХААНА ч байсан (content-ийн <a>, <iframe>, _embedded
      media, acf талбар, /wp/v2/media?parent=) рекурсиваар олно
    • Excel бол bank_financials_bot-ийн задлагчаар, PDF бол мөр бүрийн
      "шошго … тоо тоо" бүтцээр (баруун талын = тайлант үе) задална.
      Скан (зурган) PDF-д текст байхгүй тул алгасаж, мэдэгдэнэ.

Бие даан ажиллуулах (сүлжээтэй компьютер дээр):
    python3 capitron_source.py --list                 # олдсон бүх тайлан, улирал, файл
    python3 capitron_source.py --latest --out reports # сүүлийн улирлыг татаж, задлаад хадгална
    python3 capitron_source.py --dump-json resp.json  # API хариуг файлд (алдаа шалгахад)
    python3 capitron_source.py --from-json resp.json --list   # сүлжээгүй, хадгалсан хариугаар
Pipeline (quarterly_pipeline.py) үүнийг автоматаар дууддаг.
"""
from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urljoin

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402

BANK = "Капитрон"
TICKER = "CAPITRON"
SITE = "https://www.capitronbank.mn"
API = SITE + "/admin//wp-json/wp/v2"
CATEGORIES = (43, 44)
HEADERS = {
    "accept": "application/json, */*",
    "accept-language": "mn,en-US;q=0.9,en;q=0.8",
    "referer": SITE + "/about/20?lang=&type=about",
    "user-agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36"),
}
FILE_RE = re.compile(r"https?://[^\s\"'<>()]+?\.(?:pdf|xlsx|xls)(?:\?[^\s\"'<>]*)?", re.I)
REL_FILE_RE = re.compile(r"""(?:href|src|data)\s*=\s*["']([^"']+?\.(?:pdf|xlsx|xls)(?:\?[^"']*)?)["']""", re.I)


# ── HTTP ──────────────────────────────────────────────────────────────────
def http_get(url: str, timeout: int = 60):
    r = bot._SESSION.get(url, headers=HEADERS, timeout=timeout)
    if r.status_code >= 400 and "/admin//" in url:
        # Зарим сервер давхар "//"-г хүлээж авдаггүй — нэг "/"-тэйгээр дахин оролдоно
        r = bot._SESSION.get(url.replace("/admin//", "/admin/"), headers=HEADERS, timeout=timeout)
    r.raise_for_status()
    return r


def fetch_posts(category: int) -> list[dict]:
    """Тухайн ангиллын БҮХ нийтлэл (100-аас олон бол хуудаслана)."""
    posts, page = [], 1
    while True:
        r = http_get(f"{API}/posts?_embed=true&categories={category}&per_page=100&page={page}")
        batch = r.json()
        if not isinstance(batch, list):
            break
        for p in batch:
            p["_category"] = category
        posts += batch
        total_pages = int(r.headers.get("X-WP-TotalPages", "1") or 1)
        if page >= total_pages or not batch:
            break
        page += 1
    return posts


def fetch_media_children(post_id: int) -> list[str]:
    try:
        r = http_get(f"{API}/media?parent={post_id}&per_page=100")
        return [m.get("source_url", "") for m in r.json() if isinstance(m, dict)]
    except Exception:  # noqa: BLE001
        return []


# ── Нийтлэлээс файл, улирал олох ──────────────────────────────────────────
def _walk_strings(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from _walk_strings(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _walk_strings(v)


def post_files(post: dict) -> list[str]:
    """Нийтлэлийн аль ч хэсэгт (content, excerpt, acf, _embedded …) байгаа файлын URL."""
    found = []
    for s in _walk_strings(post):
        s = html.unescape(s).replace("\\/", "/")
        for u in FILE_RE.findall(s):
            found.append(u)
        for u in REL_FILE_RE.findall(s):
            found.append(urljoin(SITE + "/", u))
    out = []
    for u in found:
        u = u.strip()
        if u not in out:
            out.append(u)
    return out


def post_text(post: dict) -> str:
    def rendered(k):
        v = post.get(k)
        return v.get("rendered", "") if isinstance(v, dict) else str(v or "")
    t = " ".join([rendered("title"), rendered("excerpt"), rendered("content")])
    t = re.sub(r"<[^>]+>", " ", html.unescape(t))
    return re.sub(r"\s+", " ", t).strip()


ROMAN = {"i": 1, "ii": 2, "iii": 3, "iv": 4}
QEND = {(3, 31): 1, (6, 30): 2, (9, 30): 3, (12, 31): 4}


def detect_period(*texts: str) -> tuple[int, int] | None:
    """Текстээс (он, улирал). Хамгийн тодорхой хэлбэрээс эхэлнэ."""
    t = " ".join(unquote(x) for x in texts if x)
    t = re.sub(r"[_\-]+", " ", t)
    tl = t.lower()
    m = re.search(r"(20\d\d)\s*[./]\s*(\d{1,2})\s*[./]\s*(\d{1,2})", t)
    if m and (int(m.group(2)), int(m.group(3))) in QEND:
        return int(m.group(1)), QEND[(int(m.group(2)), int(m.group(3)))]
    m = re.search(r"(20\d\d)\s*оны\s*(\d{1,2})\s*(?:дугаар|дүгээр|-?р)?\s*сарын\s*(\d{1,2})", tl)
    if m and (int(m.group(2)), int(m.group(3))) in QEND:
        return int(m.group(1)), QEND[(int(m.group(2)), int(m.group(3)))]
    m = re.search(r"(20\d\d)\s*(?:оны|он)?\s*(iv|iii|ii|i)\s*(?:улирал|улирлын)", tl)
    if m:
        return int(m.group(1)), ROMAN[m.group(2)]
    m = re.search(r"(20\d\d)\s*(?:оны|он)?\s*([1-4])\s*(?:-?р|дугаар|дүгээр)?\s*(?:улирал|улирлын)", tl)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"\bq([1-4])\s*(20\d\d)\b|\b(20\d\d)\s*q([1-4])\b", tl)
    if m:
        return (int(m.group(2)), int(m.group(1))) if m.group(1) else (int(m.group(3)), int(m.group(4)))
    y = re.search(r"(20\d\d)", t)
    if y:
        year = int(y.group(1))
        if re.search(r"хагас\s*жил|h1\b", tl):
            return year, 2
        if re.search(r"9\s*сар", tl):
            return year, 3
        if re.search(r"жилийн\s*(?:эцсийн)?|аудитлагдсан|санхүүгийн\s*жил|annual", tl):
            return year, 4
    return None


def list_reports(posts: list[dict] | None = None, with_media: bool = True) -> list[dict]:
    """[{period, title, date, category, files, audited, post_id}] — шинэ нь эхэндээ."""
    if posts is None:
        posts = []
        for c in CATEGORIES:
            try:
                posts += fetch_posts(c)
            except Exception as e:  # noqa: BLE001
                print(f"  ⚠ Капитрон ангилал {c}: {e}")
    out = []
    for p in posts:
        files = post_files(p)
        if with_media and not files and p.get("id"):
            files = [u for u in fetch_media_children(p["id"]) if re.search(r"\.(pdf|xlsx?)($|\?)", u, re.I)]
        text = post_text(p)
        title = re.sub(r"<[^>]+>", "", html.unescape((p.get("title") or {}).get("rendered", "")))
        per_file = {}
        for f in files:
            per_file[f] = detect_period(Path(f.split("?")[0]).name) or detect_period(title, text)
        period = detect_period(title) or detect_period(text) or next((v for v in per_file.values() if v), None)
        out.append({"post_id": p.get("id"), "title": title.strip(), "date": p.get("date"),
                    "category": p.get("_category"), "period": period, "files": per_file,
                    "audited": bool(re.search(r"аудит", title + " " + text, re.I)), "link": p.get("link")})
    out.sort(key=lambda r: (r["period"] or (0, 0), r["date"] or ""), reverse=True)
    return out


# ── Файл задлах ───────────────────────────────────────────────────────────
NUM_TOKEN = r"\(?-?\d{1,3}(?:[ ,]\d{3})*(?:\.\d+)?\)?|\(?-?\d+(?:\.\d+)?\)?|-"
LINE_RE = re.compile(rf"^\s*(?:[\d.]+\s+)?(?P<label>[^\d()\-][^\d]*?[А-Яа-яӨөҮүЁё)][^\d]*?)\s+(?P<nums>(?:(?:{NUM_TOKEN})\s*)+)$")


def _split_nums(s: str) -> list[float | None]:
    toks = re.findall(NUM_TOKEN, s)
    vals = []
    for t in toks:
        vals.append(bot.parse_numeric_cell(t.replace(" ", "")))
    return vals


def parse_pdf(path: Path) -> dict[str, float]:
    """Текст давхаргатай PDF-ийн мөр бүрээс 'шошго … тоо тоо' -> баруун талын тоо.
    (Монголбанкны маягтад баруун багана нь тайлант үеийн үлдэгдэл/дүн байдаг —
    Excel задлагчтай ижил дүрэм.)"""
    import pdfplumber
    raw: dict[str, float] = {}
    seen = set()
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            # Хүснэгтийг эхлээд оролдоно (багана тодорхой бол хамгийн найдвартай)
            for tbl in page.extract_tables() or []:
                for row in tbl:
                    cells = [c for c in row if c not in (None, "")]
                    lab = next((c for c in cells if isinstance(c, str) and bot.CYRILLIC_RE.search(c) and len(c.strip()) > 3), None)
                    if not lab:
                        continue
                    nums = [bot.parse_numeric_cell(str(c).replace(" ", "")) for c in cells[cells.index(lab) + 1:]]
                    nums = [n for n in nums if n is not None]
                    _add(raw, seen, lab, nums)
            for line in (page.extract_text() or "").splitlines():
                m = LINE_RE.match(line.strip())
                if not m:
                    continue
                nums = [n for n in _split_nums(m.group("nums")) if n is not None]
                _add(raw, seen, m.group("label"), nums)
    bot.aggregate_npl_categories(raw)
    return raw


def _add(raw, seen, label, nums):
    if not nums:
        return
    label = re.sub(r"\s+", " ", str(label).replace("\n", " ")).strip(" .:")
    n = bot.normalize_row_label(label)
    if n in bot.SECTION_HEADER_LABELS or n in seen or len(n) < 4:
        return
    seen.add(n)
    raw[label] = nums[-1]


def parse_file(path: Path) -> dict[str, float]:
    if path.suffix.lower() == ".pdf":
        return parse_pdf(path)
    import pandas as pd
    raw: dict[str, float] = {}
    xls = pd.ExcelFile(path)
    for sh in xls.sheet_names:
        bot.extract_sheet_labels(pd.read_excel(xls, sheet_name=sh, header=None), raw)
    bot.aggregate_npl_categories(raw)
    return raw


def adapt_to_sheet(raw: dict[str, float]) -> dict[str, float]:
    """Капитрон sheet-ийн өвөрмөц мөрүүдэд (master-ийг аудитын тайлангаас
    бүтээсэн тул стандарт маягтаас өөр нэртэй) эх тайлангийн мөрийг нэгтгэнэ.
    Аль хэдийн байгаа шошгыг дарахгүй."""
    out = dict(raw)
    low = {bot.normalize_row_label(k): v for k, v in raw.items()}

    def get(*names):
        for n in names:
            if n in low:
                return low[n]
        return None

    def put(label, val):
        if val is not None and bot.normalize_row_label(label) not in low:
            out[label] = val

    cur, dep = get("харилцах"), get("хадгаламж")
    if cur is not None and dep is not None:
        put("Харилцагчдын харилцах, хадгаламж", cur + dep)
    oth = [get("бусад санхүүгийн хөрөнгө"), get("бусад санхүүгийн бус хөрөнгө")]
    if any(v is not None for v in oth):
        put("Бусад хөрөнгө", sum(v for v in oth if v is not None))
    put("Хөрөнгө оруулалт (санхүүгийн)", get("хөрөнгө оруулалт"))
    ol = [get("бусад өр төлбөр"), get("бусад санхүүгийн өр төлбөр"), get("бусад санхүүгийн бус өр төлбөр")]
    if ol[0] is not None:
        put("Бусад өр төлбөр (репо, бусад)", ol[0])
    elif any(v is not None for v in ol[1:]):
        put("Бусад өр төлбөр (репо, бусад)", sum(v for v in ol[1:] if v is not None))
    put("Хуримтлагдсан ашиг, алдагдал", get("хуримтлагдсан ашиг", "хуримтлагдсан ашиг (алдагдал)"))
    put("ӨӨРИЙН ХӨРӨНГИЙН ДҮН", get("нийт өөрийн хөрөнгө", "өөрийн хөрөнгийн дүн"))
    put("НИЙТ ӨР ТӨЛБӨРИЙН ДҮН", get("нийт өр төлбөр", "өр төлбөрийн дүн"))
    put("НИЙТ ХӨРӨНГИЙН ДҮН", get("нийт хөрөнгө", "хөрөнгийн дүн"))
    return out


# ── Pipeline-д зориулсан нэр дэвшигч ──────────────────────────────────────
def candidates_for(min_period, reports_dir: Path, reports: list[dict] | None = None) -> tuple[tuple | None, list[dict], str]:
    """Сүүлийн улирлын (≥ min_period) файлуудыг татаж задлаад pipeline-ийн
    нэр дэвшигч хэлбэрээр буцаана: (period, candidates, heading)."""
    reports = reports if reports is not None else list_reports()
    dated = [r for r in reports if r["period"] and r["files"]]
    if not dated:
        return None, [], "тайлан олдсонгүй"
    top = max(r["period"] for r in dated)
    heading = " | ".join(r["title"] for r in dated if r["period"] == top)
    if top < min_period:
        return top, [], heading
    cands = []
    for r in dated:
        if r["period"] != top:
            continue
        for url, fper in r["files"].items():
            if fper and fper != top:
                continue
            try:
                resp = http_get(url)
            except Exception as e:  # noqa: BLE001
                print(f"  ⚠ Капитрон файл татсангүй {url}: {e}")
                continue
            ext = Path(url.split("?")[0]).suffix.lower() or ".pdf"
            sha = hashlib.sha256(resp.content).hexdigest()
            d = reports_dir / f"{top[0]}Q{top[1]}"
            d.mkdir(parents=True, exist_ok=True)
            path = d / f"{TICKER}_{sha[:10]}{ext}"
            if not path.exists():
                path.write_bytes(resp.content)
            try:
                raw = parse_file(path)
            except Exception as e:  # noqa: BLE001
                print(f"  ⚠ Капитрон {path.name} задлагдсангүй: {e}")
                continue
            if len(raw) < 8:
                print(f"  ⚠ Капитрон {path.name}: {len(raw)} шошго — скан PDF эсвэл тайлан биш файл байж болзошгүй")
                continue
            raw = adapt_to_sheet(raw)
            print(f"  · {path.name}: {len(raw)} шошго ({'аудитлагдсан' if r['audited'] else 'улирлын'})")
            # Аудитлагдсан хувилбарт бага зэрэг давуу эрх (ижил улиралд аль аль нь байвал)
            cands.append({"url": url, "path": path, "sha": sha, "raw": raw,
                          "dates": {top}, "audited": r["audited"]})
    return top, cands, heading


# ── CLI ───────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--latest", action="store_true")
    ap.add_argument("--out", default="reports")
    ap.add_argument("--dump-json", default=None)
    ap.add_argument("--from-json", default=None)
    a = ap.parse_args()

    posts = None
    if a.from_json:
        posts = json.loads(Path(a.from_json).read_text(encoding="utf-8"))
    if a.dump_json:
        posts = []
        for c in CATEGORIES:
            posts += fetch_posts(c)
        Path(a.dump_json).write_text(json.dumps(posts, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"✓ {len(posts)} нийтлэл -> {a.dump_json}")
    reports = list_reports(posts, with_media=not a.from_json)
    if a.list or not a.latest:
        for r in reports:
            p = f"{r['period'][0]}Q{r['period'][1]}" if r["period"] else "????"
            print(f"{p}  [{r['category']}] {r['date'][:10] if r['date'] else ''}  {r['title'][:70]}"
                  f"{'  (аудит)' if r['audited'] else ''}")
            for f, fp in r["files"].items():
                print(f"        {f}" + ("" if fp in (None, r["period"]) else f"   ← файлын нэрээр {fp}"))
    if a.latest:
        top, cands, heading = candidates_for((0, 0), Path(a.out), reports)
        print(f"\nСүүлийн улирал: {top} — {heading}")
        for c in cands:
            j = c["path"].with_suffix(".labels.json")
            j.write_text(json.dumps(c["raw"], ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"  ✓ {c['path']}  ({len(c['raw'])} шошго) -> {j.name}")


if __name__ == "__main__":
    main()
