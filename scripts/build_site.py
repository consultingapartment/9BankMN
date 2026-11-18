#!/usr/bin/env python3
"""
build_site.py — data/price_data.json болон data/fundamentals.json-ыг
template/dashboard_template.html дотор шингээж, docs/index.html
(GitHub Pages-ийн эх файл) болгож бичнэ.

Ажиллуулах:
    python3 scripts/build_site.py
"""
import re
import json
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
TEMPLATE_FILE = REPO_ROOT / "template" / "dashboard_template.html"
PRICE_FILE = REPO_ROOT / "data" / "price_data.json"
FUND_FILE = REPO_ROOT / "data" / "fundamentals.json"
OUTPUT_FILE = REPO_ROOT / "docs" / "index.html"


def main():
    html = TEMPLATE_FILE.read_text(encoding="utf-8")
    price_json = PRICE_FILE.read_text(encoding="utf-8")
    fund_json = FUND_FILE.read_text(encoding="utf-8")

    # JSON бүтэн бүтэцтэй эсэхийг эхлээд шалгаад авъя (эвдэрсэн өгөгдлөөр
    # хуучин ажиллаж байсан сайтыг дарж бичихээс сэргийлнэ)
    json.loads(price_json)
    json.loads(fund_json)

    html = html.replace("/*__PRICE_JSON__*/", price_json)
    html = html.replace("/*__FUND_JSON__*/", fund_json)

    from datetime import datetime, timezone
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    html = html.replace("<!--LAST_UPDATED-->", stamp)

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_FILE.write_text(html, encoding="utf-8")
    print(f"✓ {OUTPUT_FILE} бичигдлээ ({OUTPUT_FILE.stat().st_size // 1024} KB), огноо тэмдэглэл: {stamp}")


if __name__ == "__main__":
    main()
