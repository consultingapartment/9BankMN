#!/usr/bin/env python3
"""
transbank_source.py — Транс банкны улирлын санхүүгийн тайланг transbank.mn (Next.js)
сайтаас олж, татаж, задална.

Сайт Next.js тул тайлангийн жагсаалт нь хуудасны JSON (pageProps)-д байдаг:
    https://transbank.mn/_next/data/<buildId>/mn/about-us/financial-information/quarterly-financial-statements.json
<buildId> (жиш. BzpLK5GdUFXH15MmVUYUp) сайт шинэчлэгдэх БҮРТ өөрчлөгддөг тул
кодонд хатуу бичихгүй. Оронд нь:
    1. HTML хуудсыг татаж <script id="__NEXT_DATA__">-аас pageProps-ийг шууд авна
       (тэнд buildId ч бас бий);
    2. HTML-д pageProps хоосон бол buildId-аар _next/data JSON-ийг татна;
    3. Аль нь ч болохгүй бол сүүлд мэдэгдэж байсан buildId-г оролдоно.

pageProps-ийн бүтцийг урьдчилан мэдэхгүй тул (CMS-ээс хамаарна) JSON-ийг бүхэлд нь
рекурсиваар гүйж, .pdf/.xlsx/.xls файл бүрийг ОЙРЫН эцэг объектын текстээр (гарчиг,
он, улирал талбар) нь улирлыг тодорхойлно.

    python3 transbank_source.py --list
    python3 transbank_source.py --latest --out reports
    python3 transbank_source.py --dump-json transbank.json      # pageProps-ийг хадгалах
    python3 transbank_source.py --from-json transbank.json --list
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urljoin

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import bank_financials_bot as bot  # noqa: E402
import capitron_source as common  # noqa: E402  (detect_period, parse_file — ерөнхий туслахууд)

BANK = "Транс банк"
TICKER = "TRANS"
SITE = "https://transbank.mn"
PATH = "about-us/financial-information/quarterly-financial-statements"
PAGES = [f"{SITE}/mn/{PATH}", f"{SITE}/{PATH}"]
KNOWN_BUILD_ID = "BzpLK5GdUFXH15MmVUYUp"        # 2026/11-д ажиглагдсан — зөвхөн нөөц
HEADERS = {
    "accept": "text/html,application/json;q=0.9,*/*;q=0.8",
    "accept-language": "mn,en-US;q=0.9,en;q=0.8",
    "referer": f"{SITE}/{PATH}",
    "user-agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36"),
}
FILE_EXT_RE = re.compile(r"\.(pdf|xlsx|xls)(?:$|[?#])", re.I)
NEXT_DATA_RE = re.compile(r'<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', re.S)


def http_get(url: str, timeout: int = 60, extra: dict | None = None):
    r = bot._SESSION.get(url, headers={**HEADERS, **(extra or {})}, timeout=timeout)
    r.raise_for_status()
    return r


# ── pageProps авах ────────────────────────────────────────────────────────
def fetch_page_props() -> tuple[dict, str | None]:
    """(pageProps, buildId)."""
    build_id, last_err = None, None
    for url in PAGES:
        try:
            html_text = http_get(url).text
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
        m = NEXT_DATA_RE.search(html_text)
        if not m:
            continue
        nd = json.loads(m.group(1))
        build_id = nd.get("buildId") or build_id
        props = (nd.get("props") or {}).get("pageProps") or {}
        if props and _has_files(props):
            return props, build_id
    for bid in [b for b in (build_id, KNOWN_BUILD_ID) if b]:
        for loc in ("mn/", ""):
            url = f"{SITE}/_next/data/{bid}/{loc}{PATH}.json"
            try:
                js = http_get(url, extra={"x-nextjs-data": "1", "accept": "application/json"}).json()
            except Exception as e:  # noqa: BLE001
                last_err = e
                continue
            props = js.get("pageProps", js)
            if props:
                return props, bid
    raise RuntimeError(f"Транс банкны pageProps олдсонгүй ({last_err})")


# ── JSON-оос файл + улирал ────────────────────────────────────────────────
def _scalars(d: dict) -> list[str]:
    out = []
    for k, v in d.items():
        if isinstance(v, (str, int, float)) and not isinstance(v, bool):
            sv = str(v)
            if re.match(r"\d{4}-\d{2}-\d{2}T", sv):   # createdAt/updatedAt — улирал биш
                continue
            if not FILE_EXT_RE.search(sv) and len(sv) < 300:
                out.append(f"{k}: {sv}")
    return out


def _explicit_period(d: dict) -> tuple[int, int] | None:
    """{"year": 2026, "quarter": 3} / {"он": 2026, "улирал": "III"} гэх мэт талбар."""
    y = q = None
    for k, v in d.items():
        kl = str(k).lower()
        if re.fullmatch(r"(year|он|jil|жил)", kl):
            try:
                y = int(str(v)[:4])
            except ValueError:
                pass
        elif re.fullmatch(r"(quarter|season|улирал|uliral|q)", kl):
            s = str(v).strip().lower()
            q = common.ROMAN.get(s) or (int(s[0]) if s[:1].isdigit() and 1 <= int(s[0]) <= 4 else None)
    return (y, q) if y and q and 2000 < y < 2100 else None


def _file_url(d: dict) -> str | None:
    """Strapi/CMS-ийн {url: '/uploads/x', ext: '.pdf'} эсвэл mime-тэй объект."""
    url = d.get("url") or d.get("src") or d.get("href")
    if not isinstance(url, str):
        return None
    ext = str(d.get("ext") or "").lower()
    mime = str(d.get("mime") or d.get("mimeType") or "").lower()
    if FILE_EXT_RE.search(url) or ext in (".pdf", ".xlsx", ".xls") or "pdf" in mime or "spreadsheet" in mime or "excel" in mime:
        return url
    return None


def extract_files(props) -> list[dict]:
    """[{url, ctx:[дотроос гадагш текст түвшнүүд], explicit:(он,улирал)|None}]"""
    found: list[dict] = []
    seen = set()

    def emit(url, ctx, explicit):
        url = url.strip()
        if url.startswith("//"):
            url = "https:" + url
        elif not url.startswith("http"):
            url = urljoin(SITE + "/", url.lstrip("/"))
        if url in seen:
            return
        seen.add(url)
        found.append({"url": url, "ctx": list(ctx), "explicit": explicit})

    def walk(obj, ctx, explicit):
        if isinstance(obj, dict):
            ctx2 = [" | ".join(_scalars(obj))] + ctx
            exp2 = _explicit_period(obj) or explicit
            u = _file_url(obj)
            if u:
                emit(u, ctx2, exp2)
            for v in obj.values():
                if isinstance(v, str):
                    if FILE_EXT_RE.search(v) and ("/" in v or v.startswith("http")):
                        emit(v, ctx2, exp2)
                    for href in re.findall(r"""href=["']([^"']+\.(?:pdf|xlsx|xls)[^"']*)["']""", v, re.I):
                        emit(href, ctx2, exp2)
                else:
                    walk(v, ctx2, exp2)
        elif isinstance(obj, list):
            for v in obj:
                walk(v, ctx, explicit)
    walk(props, [], None)
    return found


def list_reports(props: dict | None = None) -> list[dict]:
    props = props if props is not None else fetch_page_props()[0]
    out = []
    for f in extract_files(props):
        name = unquote(Path(f["url"].split("?")[0]).name)
        period = common.detect_period(name)
        if not period:
            for level in f["ctx"]:                    # хамгийн ойрын эцгээс эхэлнэ
                period = common.detect_period(level)
                if period:
                    break
        period = period or f["explicit"]
        title = next((lvl for lvl in f["ctx"] if lvl), name)[:120]
        out.append({"url": f["url"], "name": name, "period": period, "title": title,
                    "audited": bool(re.search(r"аудит|audit", " ".join(f["ctx"][:2]) + name, re.I))})
    out.sort(key=lambda r: r["period"] or (0, 0), reverse=True)
    return out


def _has_files(props) -> bool:
    return bool(extract_files(props))


parse_file = common.parse_file


def adapt_to_sheet(raw: dict[str, float]) -> dict[str, float]:
    """Транс банкны sheet стандарт маягтын нэртэй тул нэмэлт нэгтгэл хэрэггүй."""
    return raw


def candidates_for(min_period, reports_dir: Path, reports: list[dict] | None = None):
    reports = reports if reports is not None else list_reports()
    dated = [r for r in reports if r["period"]]
    if not dated:
        return None, [], "тайлан олдсонгүй"
    top = max(r["period"] for r in dated)
    heading = " | ".join(sorted({r["name"] for r in dated if r["period"] == top}))
    if top < min_period:
        return top, [], heading
    cands = []
    for r in dated:
        if r["period"] != top:
            continue
        try:
            resp = http_get(r["url"])
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠ Транс банк файл татсангүй {r['url']}: {e}")
            continue
        ext = Path(r["url"].split("?")[0]).suffix.lower()
        if ext not in (".pdf", ".xlsx", ".xls"):
            ext = ".pdf" if resp.content[:4] == b"%PDF" else ".xlsx"
        sha = hashlib.sha256(resp.content).hexdigest()
        d = reports_dir / f"{top[0]}Q{top[1]}"
        d.mkdir(parents=True, exist_ok=True)
        path = d / f"{TICKER}_{sha[:10]}{ext}"
        if not path.exists():
            path.write_bytes(resp.content)
        try:
            raw = adapt_to_sheet(common.parse_file(path))
        except Exception as e:  # noqa: BLE001
            print(f"  ⚠ Транс банк {path.name} задлагдсангүй: {e}")
            continue
        if len(raw) < 8:
            print(f"  ⚠ Транс банк {path.name}: {len(raw)} шошго — скан PDF эсвэл тайлан биш файл байж болзошгүй")
            continue
        print(f"  · {path.name}: {len(raw)} шошго ({r['name']})")
        cands.append({"url": r["url"], "path": path, "sha": sha, "raw": raw, "dates": {top}, "audited": r["audited"]})
    return top, cands, heading


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--latest", action="store_true")
    ap.add_argument("--out", default="reports")
    ap.add_argument("--dump-json", default=None)
    ap.add_argument("--from-json", default=None)
    a = ap.parse_args()
    props = None
    if a.from_json:
        props = json.loads(Path(a.from_json).read_text(encoding="utf-8"))
    else:
        props, bid = fetch_page_props()
        print(f"buildId: {bid}")
    if a.dump_json:
        Path(a.dump_json).write_text(json.dumps(props, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"✓ pageProps -> {a.dump_json}")
    reports = list_reports(props)
    if a.list or not a.latest:
        for r in reports:
            p = f"{r['period'][0]}Q{r['period'][1]}" if r["period"] else "????"
            print(f"{p}  {r['name'][:60]:60}  {r['title'][:60]}")
            print(f"        {r['url']}")
    if a.latest:
        top, cands, heading = candidates_for((0, 0), Path(a.out), reports)
        print(f"\nСүүлийн улирал: {top} — {heading}")
        for c in cands:
            j = c["path"].with_suffix(".labels.json")
            j.write_text(json.dumps(c["raw"], ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"  ✓ {c['path']} ({len(c['raw'])} шошго) -> {j.name}")


if __name__ == "__main__":
    main()
