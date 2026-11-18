#!/usr/bin/env python3
"""
bank_financials_bot.py — 5 банкны (SBM/564, XAC/568, TDB/567, GLMT/562, KHAN/563)
хамгийн сүүлийн улирлын тайланг open.mse.mn-ээс шалгаж, Excel-ээс шинэ бол
татаж бичээд, бусад бүх хуудсыг (Дашборд, Харьцуулалт, CAMELS г.м.) дахин
тооцоолуулдаг бот.

Урсгал (яг таны хүссэнээр):
    1. open.mse.mn/securities/{id}/tab/financials -> аль улирлын тайлан вэ?
    2. Excel файл дахь тухайн банкны сүүлийн огноо хэзээнийх вэ?
    3. Хэрэв API дээрх улирал Excel-ээс шинэ бол -> тэр улирлын БҮХ
       мэдээллийг татаж, шинэ багана болгон бичнэ
    4. LibreOffice-оор бүх томьёог дахин тооцоолуулна (Дашборд гэх мэт
       бусад хуудсууд автоматаар шинэчлэгдэнэ)
    5. Telegram-руу үр дүнгийн мэдэгдэл илгээнэ

⚠️ Claude-ийн орчинд ажиллуулж БОЛОХГҮЙ — энд open.mse.mn рүү сүлжээгээр
   хандах эрхгүй. Та өөрийн компьютер/серверт ажиллуулна уу.

Ажиллуулах:
    pip install requests beautifulsoup4 openpyxl pandas xlrd
    export TELEGRAM_BOT_TOKEN="..."      # (сонголт, өгөхгүй бол Telegram алгасна)
    export TELEGRAM_CHAT_ID="..."
    python3 bank_financials_bot.py Есөн_банк_шинжилгээ2026-06-30.xlsx
    python3 bank_financials_bot.py Есөн_банк_шинжилгээ2026-06-30.xlsx --check   # татахгүй, зөвхөн шалгана
"""
import os
import re
import sys
import subprocess
from pathlib import Path
from datetime import datetime

import requests
import pandas as pd
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
from openpyxl import load_workbook

# "Зохистой харьцаа" sheet-ийг бот шинэчлэхгүй (гараар/өөр эх сурвалжаас). True болговол
# update_ratio_sheet ажиллана.
UPDATE_RATIO_SHEET = False

BASE_URL = "https://open.mse.mn"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; BankFinancialsBot/1.0)"}

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

BANKS = {
    "Төрийн банк": {"ticker": "SBM", "id": 564},
    "ХасБанк":     {"ticker": "XAC", "id": 568},
    "ХХБ":         {"ticker": "TDB", "id": 567},
    "Голомт":      {"ticker": "GLMT", "id": 562},
    "Хаан банк":   {"ticker": "KHAN", "id": 563},
}

# Банк тус бүрийн баланс/орлого хуудасны толгой мөр (Excel-ээс баталгаажуулсан)
HEADER_ROW = {
    "Төрийн банк": {"bal": 4, "inc": 4},
    "ХасБанк":     {"bal": 4, "inc": 4},
    "ХХБ":         {"bal": 5, "inc": 4},
    "Голомт":      {"bal": 4, "inc": 4},
    "Хаан банк":   {"bal": 4, "inc": 4},
}


def fetch_quarter_info(company_id: int) -> dict | None:
    """
    /tab/financials хуудаснаас ЗӨВХӨН (year, quarter) мэдээллийг л авна.

    ⚠️ ЭНЭ ХУУДСАНААС ӨГӨГДӨЛ (тоо) АВАХГҮЙ — зөвхөн аль улирлын тайлан
    хамгийн сүүлд гарсныг мэдэхэд ашиглана. Учир нь энд байгаа
    .finance-report-item блок ердөө ~14-25 хураангуй тоотой (маркетингийн
    хураангуй түвшин) бөгөөд Master Excel-ийн 35-40 мөртэй дэлгэрэнгүй
    баланс/орлогыг ер нь гүйцээх боломжгүй. Бодит дэлгэрэнгүй өгөгдлийг
    ЗӨВХӨН fetch_full_report() (хавсаргасан албан ёсны .xls/.xlsx файл)-ээс
    л авна.
    """
    url = f"{BASE_URL}/securities/{company_id}/tab/financials"
    resp = safe_get(url)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    block = soup.select_one(".finance-report-item")
    if block is None:
        return None

    heading = block.find("h5")
    heading_text = heading.get_text(strip=True) if heading else ""
    m = re.search(r"(\d{4})\s*Он\s*(\d)\s*Улирал", heading_text)
    if not m:
        return None
    year, quarter = int(m.group(1)), int(m.group(2))
    return {"year": year, "quarter": quarter, "heading": heading_text}


def get_session() -> requests.Session:
    """open.mse.mn сервер ачаалалтай үед timeout/5xx өгдөг тул 3 удаа
    автоматаар давтан оролддог Session (Gemini-ийн санаа) — бүх татах
    дуудлага үүгээр дамжина."""
    session = requests.Session()
    retries = Retry(total=3, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    session.mount("https://", HTTPAdapter(max_retries=retries))
    session.mount("http://", HTTPAdapter(max_retries=retries))
    return session


_SESSION = get_session()


def safe_get(url: str, **kwargs):
    """Эхлээд ердийн (баталгаажуулалттай) татна; зөвхөн ЯГ SSL-ийн гинжний
    алдаа гарвал л (Mn-stock репозиторийн README-д тэмдэглэсэн зарим
    МХБ-тэй холбоотой сайт кэбийн дутуу гинжтэй байдаг тул) verify=False-
    ээр дахин оролдоно — Gemini-ийн өгсөн кодонд байсан зохистой санаа,
    гэхдээ ямар ч үед бус ЗӨВХӨН шаардлагатай үед л ашиглана. Мөн Retry
    session ашигладаг тул timeout/5xx үед автоматаар 3 удаа давтана."""
    try:
        return _SESSION.get(url, headers=HEADERS, timeout=kwargs.pop("timeout", 30), **kwargs)
    except requests.exceptions.SSLError:
        print(f"  ⚠ SSL гэрчилгээний гинж дутуу — баталгаажуулалтгүйгээр дахин оролдож байна: {url}")
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        return _SESSION.get(url, headers=HEADERS, timeout=30, verify=False, **kwargs)


def aggregate_npl_categories(raw_labels: dict[str, float]) -> None:
    """
    Эх тайлан 5 зэрэглэлтэй зээлийн ангилал ашигладаг (Хэвийн, Хугацаа
    хэтэрсэн, Хэвийн бус, Эргэлзээтэй, Муу), харин манай workbook 3
    түвшинд (Хэвийн, Анхаарал хандуулах, Чанаргүй) бүлэглэдэг. Монголбанкны
    стандарт ангиллын дагуу:
        Анхаарал хандуулах зээл = Хугацаа хэтэрсэн зээл
        Чанаргүй зээл = Хэвийн бус + Эргэлзээтэй + Муу зээл
    Бодит тайлангаас шалгаж, Төрийн банкны жинхэнэ утгатай 100% таарсныг
    баталгаажуулсан (Хугацаа хэтэрсэн=68,940.0≈Анхаарал хандуулах;
    Хэвийн бус+Эргэлзээтэй+Муу=199,337.4≈Чанаргүй зээл).
    Raw_labels-д аль хэдийн байгаа шошгыг ДАРЖ БИЧИХГҮЙ, зөвхөн байхгүй
    үед л нэмнэ.
    """
    if "Хугацаа хэтэрсэн зээл" in raw_labels and "Анхаарал хандуулах зээл" not in raw_labels:
        raw_labels["Анхаарал хандуулах зээл"] = raw_labels["Хугацаа хэтэрсэн зээл"]

    npl_parts = ["Хэвийн бус зээл", "Эргэлзээтэй зээл", "Муу зээл"]
    if all(p in raw_labels for p in npl_parts) and "Чанаргүй зээл" not in raw_labels:
        raw_labels["Чанаргүй зээл"] = sum(raw_labels[p] for p in npl_parts)


def fetch_full_report(company_id: int) -> dict | None:
    """
    ✅ ҮНДСЭН ба ЦОРЫН ГАНЦ эх сурвалж — банкны хавсаргасан АЛБАН ЁСНЫ
    .xls/.xlsx тайлан файлыг татаж, бүх (шошго -> утга) хосыг задална.

    open.mse.mn/tab/financials-ийн HTML хуудаснаас ОГТ ӨГӨГДӨЛ АВАХГҮЙ —
    зөвхөн энэ функцээр олдсон хавсралт файлыг л ашиглана, учир нь HTML
    хуудасны хураангуй (~14-25 талбар) нь Master Excel-ийн 35-40 мөртэй
    дэлгэрэнгүй баланс/орлогыг ер нь гүйцээх боломжгүй. 5 банк дээр бодит
    сүлжээгээр туршиж, БҮГДЭД нь (136-173 талбартайгаар) амжилттай
    ажилласныг баталгаажуулсан.
    """
    url = f"{BASE_URL}/securities/{company_id}/tab/financials"
    resp = safe_get(url)
    resp.raise_for_status()

    matches = re.findall(r'href=["\']([^"\']*\.xls[x]?)["\']', resp.text, re.IGNORECASE)
    if not matches:
        return None  # хавсралт файл олдсонгүй — бодит алдаа, бичих зүйлгүй

    link = matches[0]
    if not link.startswith("http"):
        link = BASE_URL + link
    print(f"  ℹ Албан ёсны тайлан файл: {link}")

    file_resp = safe_get(link, timeout=30)
    import tempfile
    suffix = ".xlsx" if link.lower().endswith(".xlsx") else ".xls"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, prefix=f"mse_{company_id}_") as tf:
        tf.write(file_resp.content)
        tmp_path = tf.name
    xls = pd.ExcelFile(tmp_path)
    raw_labels: dict[str, float] = {}
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        extract_sheet_labels(df, raw_labels)
    os.remove(tmp_path)
    aggregate_npl_categories(raw_labels)
    return {"raw_labels": raw_labels, "source": link}


CYRILLIC_RE = re.compile(r"[А-Яа-яЁёӨөҮү]")


def parse_numeric_cell(v):
    """Тоо (int/float) эсвэл нүдэнд бичигдсэн тоон текстийг (таслалтай,
    хаалтанд сөрөг, зураас/em-dash хоосон гэх мэт) бодит float болгоно.
    Бодит файлуудаас (golomt_bank_2026-2.xlsx) '222,676', '(33,740)', '-'
    хэлбэрүүд олдсоныг харгалзсан."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if pd.notna(v) else None
    if isinstance(v, str):
        s = v.strip()
        if s in ("", "-", "—", "***", "N/A"):
            return None
        neg = s.startswith("(") and s.endswith(")")
        s2 = s.strip("()").replace(",", "").replace(" ", "").replace("%", "")
        try:
            val = float(s2)
        except ValueError:
            return None
        return -val if neg else val
    return None


def extract_sheet_labels(df, raw_labels: dict[str, float]) -> None:
    """
    Нэг sheet-ийн (DataFrame) БҮХ мөрөөс (шошго -> утга) хосыг олж,
    raw_labels-д НЭМНЭ (in-place).

    ⚠️ ЕРӨНХИЙ, БҮТЭЦЭЭС ХАРААТ БУС АЛГОРИТМ — 2026 оны 9-р сард 6 ӨӨР
    бодит файл (issuers.mse.mn-ийн 4-sheet СБД/ОДТ формат, мөн 4 банкны
    өөрсдийн сайтаас татсан 5 ондоо загвар: statebank/khan_bank/
    golomt_bank Sheet1, XacBank-ийн BS/IS/Prudential ratios) шалгаж үзэхэд
    ХАТУУ баганы индекс, тэр байтугай ТОДОРХОЙ толгой үг ("Үзүүлэлт",
    "Балансын зүйл" гэх мэт) ч файл бүрд өөр байсан тул ямар ч тогтмол
    дүрэм НАЙДВАРТАЙ ажиллахгүй нь тогтоогдсон.

    Үүний оронд бүх мөрд ЕРӨНХИЙ дүрэм хэрэглэнэ:
      1. Мөрийн ЗҮҮН талаас эхлээд анхны бодит (кирилл үсэгтэй, 3-аас
         урт) текстийг ШОШГО гэж үзнэ (энэ нь дугаарлалт '1.1'-ийг
         автоматаар алгасна — дугаарт кирилл үсэг байдаггүй; мөн англи
         орчуулга баганыг ч алгасна — латин үсэг кирилл биш).
      2. Тэр баганаас БАРУУН тийш хамгийн сүүлийн (rightmost) ямар нэгэн
         байдлаар тоо болгож болох нүдийг УТГА гэж үзнэ — энэ нь бүх
         ажигласан файлд яг "хамгийн сүүлийн үеийн (Эцсийн үлдэгдэл)"
         баганатай тохирдог (файл бүр өөр баганын тоотой ч, СҮҮЛИЙН
         багана нь тогтмол хамгийн сүүлийн үеийн утга байдаг).
    Хэсгийн гарчиг мөр (утгагүй, жиш. "ХӨРӨНГӨ") мөр бүрт УТГА олдохгүй
    тул байгалиараа алгасагдана нэмэлт шалгуургүйгээр.
    """
    n_rows, n_cols = df.shape
    seen_norm = {normalize_row_label(k) for k in raw_labels}
    for r in range(n_rows):
        row = df.iloc[r]
        label, label_col = None, None
        for c in range(n_cols):
            v = row.iloc[c]
            if isinstance(v, str):
                text = v.strip()
                if len(text) > 3 and CYRILLIC_RE.search(text):
                    label = re.sub(r"\s+", " ", text.replace("\n", " ")).strip()
                    label_col = c
                    break
        if label is None:
            continue

        value = None
        for c in range(n_cols - 1, label_col, -1):
            val = parse_numeric_cell(row.iloc[c])
            if val is not None:
                value = val
                break
        if value is None:
            continue

        # ЧУХАЛ: тайлангийн ЭХНИЙ давталтыг хадгална, ДАРААХИЙГ нь
        # үл тоомсорлоно. Учир нь ижил шошго (жиш. "Зээлийн эрсдэлийн
        # сан") ихэвчлэн үндсэн тайланд НЭГ удаа, дараа нь тэмдэглэл/
        # задаргаа хэсэгт (рollforward schedule) ДАХИН гарч ирдэг бөгөөд
        # тэнд ихэвчлэн өөр утгатай байдаг. (Нэмэлт хамгаалалт:
        # write_new_quarter_full-ийн plausibility_score функц ижил мөр
        # рүү хэд хэдэн шошго тохирвол ч хамгийн итгэл үзүүлэхээрхийг
        # нь сонгодог.)
        # Хэсгийн гарчиг мөр (жиш. "ӨӨРИЙН ХӨРӨНГӨ") нь санамсаргүй 0 утгатай
        # байж, "НИЙТ ӨӨРИЙН ХӨРӨНГӨ" мөрийг substring-ээр буруу эзэлдэг
        # байсныг ХасБанкны Q3 файл дээр олсон — эх сурвалжаас шууд хасна.
        norm_label = normalize_row_label(label)
        if norm_label in SECTION_HEADER_LABELS:
            continue
        # Ижил НОРМЧИЛСОН нэртэй мөр (жиш. "Тайлант үеийн цэвэр ашиг, алдагдал
        # (10+11)" ба дагавар/томьёогүй хувилбар) өөр sheet-д (мөнгөн гүйлгээ,
        # тэмдэглэл) дахин гарвал ЭХНИЙ (үндсэн тайлангийн) мөрийг л авна —
        # эс бөгөөс тусдаа түлхүүр болж ХХБ-д 71,237 vs 200,550 андуурал гарна.
        if norm_label in seen_norm:
            continue
        seen_norm.add(norm_label)
        if label not in raw_labels:
            raw_labels[label] = value


GOLOMT_SITE_URL = "https://www.golomtbank.com/investor-relations/season-report"
_GOLOMT_HDR = "Аудитлагдаагүй санхүүгийн хураангуй тайлан"
_GOLOMT_RATIOS = [
    ("Нэгдүгээр зэрэглэлийн өөрийн хөрөнгө болон эрсдэлээр жигнэсэн активын зохистой харьцаа",
     r"Нэгдүгээр зэрэглэлийн өөрийн хөрөнгө болон эрсдэлээр жигнэсэн активын зохистой харьцаа.*?\s(\d+\.\d+)%"),
    ("Өөрийн хөрөнгө болон эрсдэлээр жигнэсэн активын харьцаа",
     r"Өөрийн хөрөнгө болон эрсдэлээр жигнэсэн активын зохистой харьцаа.*?\s(\d+\.\d+)%"),
    ("Төлбөр түргэн гүйцэтгэх чадварын үзүүлэлт",
     r"Төлбөр түргэн гүйцэтгэх чадварын үзүүлэлт.*?\s(\d+\.\d+)%"),
    ("Гадаад валютын нээлттэй позици болон өөрийн хөрөнгийн харьцаа (нэгж валютаар)",
     r"Гадаад валютын нээлттэй позици болон өөрийн хөрөнгийн харьцаа \(нэгж валютаар\).*?\s(\d+\.\d+)%"),
    ("Гадаад валютын нээлттэй позици болон өөрийн хөрөнгийн харьцаа (нийт)",
     r"Гадаад валютын нээлттэй позици болон өөрийн хөрөнгийн харьцаа \(нийт\).*?\s(\d+\.\d+)%"),
]
_GOLOMT_ITEM_RE = re.compile(
    r"([А-Яа-яӨөҮүЁё][^()\d]*?(?:\([^)\d]*\))?)\s+(\(?-?\d[\d,]*(?:\.\d+)?\)?|-)(?=\s|$)")


def golomt_site_dates(text: str) -> list[str]:
    """Голомт банкны IR хуудсанд байгаа бүх улирлын огноо (YYYY/MM/DD), хамгийн сүүлийнх нь хамгийн их."""
    t = re.sub(r"\s+", " ", text)
    return sorted(set(re.findall(rf"{_GOLOMT_HDR} (\d{{4}}/\d{{2}}/\d{{2}})", t)))


def parse_golomt_site_text(text: str, date_label: str) -> dict[str, float]:
    """
    golomtbank.com/investor-relations/season-report хуудасны текстээс тухайн
    огнооны (YYYY/MM/DD) баланс + орлого + зохистой харьцааг (шошго -> утга) болгоно.
    Утга сая ₮ (хуудас сая төгрөгөөр), сөрөг нь хаалттай "(34,980)". Харьцаа %-ээр
    (15.72) — update_ratio_sheet 0.1572 болгож хөрвүүлнэ. Банкны ӨӨРИЙН
    (master-ийн Голомт sheet-ийг бүтээсэн) загвар тул "Хэвийн зээл" нь
    хугацаа хэтэрсэнийг агуулсан — нэмэлт нэгтгэл хэрэггүй.
    """
    t = re.sub(r"\s+", " ", text)
    start = t.find(f"{_GOLOMT_HDR} {date_label}")
    if start < 0:
        return {}
    nxt = t.find(_GOLOMT_HDR, start + 10)
    block = t[start: nxt if nxt > 0 else None]
    out: dict[str, float] = {}
    for chunk in re.split(r"\s(?=\d+(?:\.\d+)*\.?\s+[А-Яа-яӨөҮүЁё])", block):
        chunk = re.sub(r"^\d+(?:\.\d+)*\.?\s+", "", chunk.strip())
        if "зохистой харьцааны шалгуур" in chunk or "Нэгдүгээр зэрэглэлийн" in chunk:
            continue
        for lab, val in _GOLOMT_ITEM_RE.findall(chunk):
            lab = lab.strip(" ,")
            v = parse_numeric_cell(val)
            if v is None or len(lab) < 4 or lab in out:
                continue
            out[lab] = v
    for key, pat in _GOLOMT_RATIOS:
        m = re.search(pat, block)
        if m:
            out[key] = float(m.group(1))
    return out


def fetch_golomt_site(date_label: str | None = None) -> tuple[str, dict[str, float]] | None:
    """Голомтын IR хуудаснаас (date_label эсвэл хамгийн сүүлийн) улирлын өгөгдлийг татна."""
    from bs4 import BeautifulSoup as _BS
    resp = safe_get(GOLOMT_SITE_URL)
    resp.raise_for_status()
    text = _BS(resp.text, "html.parser").get_text(" ", strip=True)
    dates = golomt_site_dates(text)
    if not dates:
        return None
    d = date_label or dates[-1]
    data = parse_golomt_site_text(text, d)
    return (d, data) if data else None


def quarter_end_date(year: int, quarter: int) -> str:
    """Улирлын дугаараас Excel-ийн огнооны багана нэр (YYYY/MM/DD, 4-р улирал '*') үүсгэнэ."""
    month_day = {1: "03/31", 2: "06/30", 3: "09/30", 4: "12/31"}[quarter]
    suffix = "*" if quarter == 4 else ""
    return f"{year}/{month_day}{suffix}"


# ── 2. Excel файл дахь сүүлийн огноог олох ──────────────────────────────
def get_excel_last_date(ws, header_row: int) -> tuple[str, int]:
    """(сүүлийн огнооны текст, тэр баганы индекс) буцаана."""
    max_col = ws.max_column
    last_val = ws.cell(header_row, max_col).value
    return (str(last_val).strip() if last_val else ""), max_col


def infer_scale_factor(new_val: float, prev_val: float | None, tolerance_mult: float = 1.0) -> float | None:
    """
    Шинэ утгыг өмнөх улирлын ЯГ ТУХАЙН мөрийн утгатай харьцуулж, хэдэн
    1000-аар хуваавал хэвийн өсөлт/бууралтын хязгаарт (0.1x–8x, эсвэл
    tolerance_mult-аар өргөтгөсөн хязгаарт) орохыг олж, ТУХАЙН нэгжийн
    зөрүүг өөрөө тодорхойлно.

    tolerance_mult: ХасБанкны орлого (H1 хуримтлал vs Q1 дан дүн) шиг
    "хэвийн харьцуулалт биш" (өмнөх утга нь ижил конвенцоор биш) мөрд
    хязгаарыг өргөтгөхөд ашиглана (2.5-3x орчим, ойролцоогоор х2 өсөлт
    хүлээгдэж буйг тооцно).

    Хэрэв ЯМАР Ч 1/1000/1e6 масштаб тохирохгүй бол (жиш. огт өөр, буруу
    мөртэй андуурч тохируулсан тохиолдолд гардаг шиг) None буцаана —
    ингэснээр дуудаж буй тал энэ тохирлыг ЗАЙЛСХИЙЖ болно. Bogd_bank_2026-2
    загвар дээр батлагдсанаар эх сурвалж МЯНГАН ₮-өөр ирдэг болох нь
    харагдсан ч, өөр хуудсанд өөр нэгж байж болзошгүй тул ТОГТМОЛ бус,
    workbook-ийн өөрийнх нь өмнөх утгатай харьцуулж баталгаажуулна.
    """
    if prev_val is None or prev_val == 0 or new_val is None or new_val == 0:
        return 1.0
    ratio = abs(new_val) / abs(prev_val)
    for divisor in (1.0, 1_000.0, 1_000_000.0):
        if 0.1 <= (ratio / divisor) <= 8.0 * tolerance_mult:
            return divisor
    return None  # ямар ч масштаб тохирохгүй — магадгүй буруу мөр, бичихгүй байх нь дээр


def copy_cell_style(src_cell, target_cell):
    """Өмнөх баганын фонт/хүрээ/дэвсгэр/тооны формат/зэрэгцүүлэлтийг хуулна
    (Gemini-ийн санааг авав — шинэ багана хоосон, форматгүй ирэхээс сэргийлнэ)."""
    import copy as _copy
    if src_cell.has_style:
        target_cell.font = _copy.copy(src_cell.font)
        target_cell.border = _copy.copy(src_cell.border)
        target_cell.fill = _copy.copy(src_cell.fill)
        target_cell.number_format = src_cell.number_format
        target_cell.alignment = _copy.copy(src_cell.alignment)


# ── 3. Excel рүү шинэ баганыг бичих (БҮРЭН хувилбар — бүх боломжит мөр) ─
def normalize_row_label(s: str) -> str:
    """Excel-ийн мөрийн шошго болон API-ийн шошгыг харьцуулахын өмнө
    норчигжуулна: тоолол ('1.4.2 '), доод зураас, олон зайг арилгана.

    ⚠️ Бодит тайлан файл (564_finance_report_xls) шууд шалгаж баталгаажуулсан
    2 нарийн орлуулга нэмсэн:
      - "СБ" (Санхүүгийн байгууллага)-ийн товчлол — манай sheet "СБ" гэж
        товчилдог, эх тайлан "санхүүгийн байгууллага" гэж бүтнээр бичдэг.
      - Томьёоны дагавар "(1-2)", "(5+6-7)", "(8-9)" гэх мэт — эх тайлангийн
        мөрийн нэрэнд байдаг боловч манай sheet-д байхгүй тул хасна
        (эс бөгөөс сүүлийн үг рүү наалдаж, үгийн тааруулалтыг эвддэг байсан).
    """
    s = re.sub(r"^[\d.\s]+", "", str(s))          # тэргүүний дугаарлалт ('1.4.2') хас
    s = re.sub(r"\s*\([\d+\-*/\s]+\)\s*$", "", s)  # төгсгөлийн '(1-2)', '(8-9)' томьёо хас
    s = re.sub(r"\s+", " ", s).strip().lower()
    s = s.replace(",", "").replace(";", "")  # таслал/цэг таслал арилгах — "банк," vs "банк" ялгаанаас сэргийлнэ
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace("өгөөж /roa/", "roa").replace("өгөөж /roe/", "roe")
    s = s.replace("сб-д", "санхүүгийн байгууллагад").replace("сб-аас", "санхүүгийн байгууллагаас")
    s = s.replace(" сб ", " санхүүгийн байгууллага ")
    s = s.replace("банк санхүүгийн байгууллагаас татсан", "банкнаас татсан")
    s = s.replace("банк санхүүгийн байгууллагад байршуулсан", "банк байгууллагад байршуулсан")
    s = s.replace("хо-ын", "хөрөнгө оруулалтын")
    s = s.replace("тайлант хугацааны нийт дэлгэрэнгүй орлогын дүн", "тайлант үеийн нийт орлого")
    s = s.replace("тайлант үеийн нийт дэлгэрэнгүй орлого", "тайлант үеийн нийт орлого")
    # Хэрэглэгчийн гараар баталгаажуулсан тодорхой синоним: манай sheet-ийн
    # "Хуримтлуулсан хүүгийн авлага" (зээлийн бүлэгт байрладаг) нь эх
    # тайлангийн 4 ижил төстэй ("...д хуримтлуулж тооцсон хүүгийн авлага")
    # мөрөөс яг ЗЭЭЛТЭЙ холбоотойг нь заана (мөнгөн хөрөнгө/банкинд
    # байршуулсан/үнэт цаасных биш).
    s = s.replace("зээлд хуримтлуулж тооцсон хүүгийн авлага", "хуримтлуулсан хүүгийн авлага")
    s = s.replace("орлогын татварын зардал", "орлогын татвар (нийт)")
    s = s.replace("өөрийн хөрөнгийн дүн", "нийт өөрийн хөрөнгө")
    s = s.replace("тайлант хугацааны нийт орлогын дүн", "тайлант үеийн нийт орлого")
    return s


# Хэсгийн ГАРЧИГ мөрүүд (өгөгдлийн бодит мөр биш, тоо байх ёсгүй) — эдгээрийг
# scan_sheet_labels-аас ЗААВАЛ хасна. Бодит тестээр батлагдсан: хасаагүй үед
# "ӨР ТӨЛБӨР" гэх мэт гарчиг мөрөнд буруу тааруулалтаар их наядын хиймэл
# тоо бичигдэж байсан (Төрийн банк/ХасБанк/ХХБ-д яг ийм алдаа гарсан).
SECTION_HEADER_LABELS = {
    "хөрөнгө", "өр төлбөр", "өөрийн хөрөнгө", "орлого, зардал",
    "тэнцлийн гадуурх данс", "үзүүлэлт",
}


def scan_sheet_labels(ws, header_row: int) -> dict[str, int]:
    """Тухайн хуудсанд байгаа БҮХ (норчигжсон шошго -> мөр дугаар) толгойг
    буцаана — ГЭХДЭЭ хэсгийн ГАРЧИГ мөрүүдийг (жиш. 'ӨР ТӨЛБӨР') хасна,
    учир нь тэдгээр нь өгөгдлийн мөр биш, зөвхөн бүлэглэлийн шошго."""
    labels = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip():
            norm = normalize_row_label(v)
            if norm in SECTION_HEADER_LABELS:
                continue
            labels[norm] = r
    return labels


NON_CURRENCY_HINTS = ("roa", "roe", "eps", "p/e", "зохистой харьцаа", "хувьцааны тоо",
                       "хувьцаа тоо", "ширхэг")


def _is_non_currency_label(norm: str) -> bool:
    """НАРИЙВЧЛАЛТАЙ (бүтэн үгээр) шалгана — "хувь" гэх мэт ТУС ГАНЦААРАА
    хэт өргөн (жиш. "Хувь нийлүүлсэн хөрөнгө" мөнгөн дүнтэй санамсаргүй
    давхцдаг байсан алдааг олж заcaв) шошгыг оруулахгүй."""
    return any(h in norm for h in NON_CURRENCY_HINTS)


def adapt_aggregates(raw_labels: dict, sheet_labels: dict) -> dict:
    """
    Эх тайлангийн мөрүүдийг ЭНЭ sheet-ийн бүтцэд тааруулан нэгтгэнэ (хуулбар
    буцаана). Банк бүрийн sheet өөр задралтай тул нэг ерөнхий нэгтгэл буруу:

    1) "Бусад өр төлбөр": эх тайланд дэд дүн байхгүй, 3 бүрдүүлэгчтэй
       (Дериватив + Бусад санхүүгийн + Бусад санхүүгийн бус). Sheet-д
       "Дериватив санхүүгийн өр төлбөр" ТУСДАА мөр байвал (ХХБ) деривативыг
       оруулахгүй — эс бөгөөс 61,548 сая ₮-өөр давхар тоологдож, өр
       төлбөрийн бүрдүүлэгчдийн нийлбэр нийтээс зөрдөг байсныг Q2
       regression-оор олсон. Байхгүй бол (Төрийн банк, ХасБанк) 3-ыг нийлүүлнэ.
    2) Голомт шиг товчилсон sheet (Хугацаа хэтэрсэн/Анхаарал хандуулах мөргүй,
       ганц "Чанаргүй зээл (NPL)"): master түүхэндээ Хэвийн зээл = Хэвийн +
       Хугацаа хэтэрсэн (8,633,005 = 8,482,996 + 150,009) гэж хадгалсан.
    """
    out = dict(raw_labels)
    parts = ["Бусад санхүүгийн өр төлбөр", "Бусад санхүүгийн бус өр төлбөр"]
    deriv = "Дериватив санхүүгийн өр төлбөр"
    has_other = "бусад өр төлбөр" in sheet_labels
    if has_other and "Бусад өр төлбөр" not in out and all(p in out for p in parts):
        if "дериватив санхүүгийн өр төлбөр" not in sheet_labels and deriv in out:
            parts = parts + [deriv]
        out["Бусад өр төлбөр"] = sum(out[p] for p in parts)

    no_overdue_row = ("хугацаа хэтэрсэн зээл" not in sheet_labels
                      and "анхаарал хандуулах зээл" not in sheet_labels)
    if (no_overdue_row and "чанаргүй зээл (npl)" in sheet_labels
            and "Хугацаа хэтэрсэн зээл" in out and "Хэвийн зээл" in out):
        out["Хэвийн зээл"] = out["Хэвийн зээл"] + out["Хугацаа хэтэрсэн зээл"]
    return out


def write_new_quarter_full(ws, header_row: int, new_col: int, date_label: str,
                            raw_labels: dict[str, float | None],
                            cumulative_flow: bool = False) -> tuple[list[str], list[str]]:
    """API-аас ирсэн БҮХ (шошго->утга)-г Excel-ийн БОДИТ мөрийн шошготой
    тааруулж аль болох олон мөрийг бичнэ. Хоосон орхисон (тохирол олдоогүй)
    мөрүүдийг тусад нь буцааж, "дутуу хагас" болгохгүй байхын тулд
    Telegram мэдэгдэлд тодорхой жагсаана. Өмнөх баганын формат/масштабыг
    (стиль болон нэгжийн зөрүү) хоёуланг нь баримталж бичнэ.

    НЭГДСЭН нэр дэвшигч цуглуулалт (2026 оны 9-р сард бодит 5 банкны
    файл дээр илэрсэн "уралдааны нөхцөл" (race condition) алдааг засав):
    Өмнөх хувилбар мөр бүрийг ЭХЭЛЖ ОЧСОН api_label-д шууд олгодог байсан
    тул, эх сурвалжинд ХОЛБООТОЙ боловч БУРУУ (жиш. "...байршуулсан
    хөрөнгө" гэдгийн оронд "...байршуулсан МӨНГӨ (3 хүртэлх сарын
    хугацаатай)" гэсэн тэмдэглэлийн мөр) шошго ЗӨВ шошгоос ӨМНӨ dict-д
    орж ирвэл, зөв шошгыг "түрүүлж эзэлдэг" алдаа гарч байсныг Голомт/
    Хаан банкны бодит файл дээр олов. Одоо БҮХ (мөр, нэр дэвшигч) хосыг
    ЭХЛЭЭД цуглуулж, дараа нь мөр бүрт хамгийн сайн нэр дэвшигчийг
    (яг тохирол > substring > fuzzy, дараа нь хамгийн ойрхон/итгэлтэйг)
    сонгоно — ЭХНИЙ ирсэн эсэхээс үл хамааран.
    """
    prev_col = new_col - 1
    header_cell = ws.cell(header_row, new_col, value=date_label)
    copy_cell_style(ws.cell(header_row, prev_col), header_cell)

    sheet_labels = scan_sheet_labels(ws, header_row)
    raw_labels = adapt_aggregates(raw_labels, sheet_labels)
    norm_api = {api_label: normalize_row_label(api_label) for api_label in raw_labels
                if raw_labels[api_label] is not None}

    def plausibility(row: int, api_label: str) -> tuple:
        """Тухайн мөрийн эх сурвалжийн утгыг өмнөх баганатай харьцуулж,
        хэр 'итгэл үзүүлэхээр' байгааг оноогоор үнэлнэ (том нь дээр)."""
        val = raw_labels[api_label]
        norm = norm_api[api_label]
        if _is_non_currency_label(norm):
            return (1, 0.0)
        prev_val = ws.cell(row, prev_col).value
        prev_val = prev_val if isinstance(prev_val, (int, float)) else None
        if val == 0:
            return (0, 0.0)
        if prev_val in (None, 0):
            # Өмнөх утга хоосон — харьцуулах суурь байхгүй. "Өмнөх=0-тэй хамгийн
            # ойр" гэж үзвэл хамгийн БАГА тоотой (ихэвчлэн буруу, жиш. тэмдэглэлийн)
            # мөр түрүүлдэг байсан (Хаан банк: 6,544 vs зөв 2,423,790). Иймд
            # төвийг сахи — шошгоны ойролцоо байдал (-илүү үг), дараа нь эх
            # тайланд эхэлж гарсан мөр шийднэ.
            return (2, 0.0)
        if (val > 0) != (prev_val > 0):
            return (0, 0.0)  # тэмдэг зөрчилдөж байна — итгэхгүй
        factor = infer_scale_factor(val, prev_val, tolerance_mult=3.0 if cumulative_flow else 1.0)
        if factor is None:
            return (0, 0.0)
        return (2, -abs(val / factor - (prev_val or 0)))

    # ── БҮХ нэр дэвшигчийг (мөр -> [(api_label, tier, extra_words), ...]) цуглуулах ──
    # tier: 2=яг тохирол, 1=substring, 0=fuzzy (үгийн багц давхцал >=0.85)
    candidates_by_row: dict[int, list[tuple[str, int, int]]] = {}

    for api_label, norm in norm_api.items():
        row = sheet_labels.get(norm)
        if row is not None:
            candidates_by_row.setdefault(row, []).append((api_label, 2, len(norm.split())))
            continue  # яг тохирол олдвол цаашид substring/fuzzy хайх шаардлагагүй

        for sheet_norm, r in sheet_labels.items():
            if norm and (norm in sheet_norm or sheet_norm in norm) and len(norm) > 4:
                candidates_by_row.setdefault(r, []).append((api_label, 1, len(norm.split())))

        api_words = set(norm.split())
        if len(api_words) >= 2:  # ганц үгтэй шошгоор fuzzy хайхгүй (хэт эрсдэлтэй)
            for sheet_norm, r in sheet_labels.items():
                sheet_words = set(sheet_norm.split())
                if len(sheet_words) < 2:
                    continue
                overlap_words = api_words & sheet_words
                if len(overlap_words) < 2:
                    continue
                overlap = len(overlap_words) / len(sheet_words)
                if overlap >= 0.85:
                    candidates_by_row.setdefault(r, []).append((api_label, 0, len(norm.split())))

    # ── Мөр бүрт хамгийн сайн нэр дэвшигчийг сонгох: tier өндөр байх нь
    # хамгийн чухал, дараа нь ТААРАЛТЫН ЧАНАР (plausibility), эцэст нь
    # ХАМГИЙН ОЙРХОН (бага илүүдэл үгтэй) шошгыг илүүд үзнэ — учир нь
    # ижил төстэй боловч илүү УРТ (жиш. тэмдэглэлийн дэд ангилал) шошго
    # санамсаргүй түрүүлж болзошгүйг бодит файл дээр олсон.
    row_for_label: dict[str, int] = {}
    for row, cands in candidates_by_row.items():
        # Эрэмбэ: tier > утгын найдвартай байдлын АНГИЛАЛ (0/1/2) > шошгоны
        # ойролцоо байдал (бага илүү үг) > утгын ойролцоо байдал. Өмнө нь утгын
        # ойролцоо байдал шошгоноос түрүүлж, "...байршуулсан мөнгө (3 хүртэлх
        # сарын хугацаатай)" (тэмдэглэл) нь зөв "...байршуулсан хөрөнгө"-өөс
        # өмнөх улирлын утгад ойр гэсэн шалтгаанаар сонгогддог байсан (Хаан банк).
        best = max(cands, key=lambda c: (c[1], plausibility(row, c[0])[0], -c[2], plausibility(row, c[0])[1]))
        row_for_label[best[0]] = row

    # Файлын НИЙТ нэгж (₮ / мянган ₮ / сая ₮): өмнөх утгатай мөрүүдээс олсон
    # масштабуудын давамгайлах (mode) утга. Өмнөх улирал хоосон (None/0) мөрд
    # (жиш. Хаан банкны "Хэвийн бус зээл") мөрийн өөрийн масштабыг тодорхойлох
    # боломжгүй тул энэ файлын нийт нэгжийг хэрэглэнэ — эс бөгөөс ₮-өөр ирсэн
    # 119,759,191,173 нь сая болгож хөрвүүлэгдэлгүй бичигдэж байсан.
    from collections import Counter
    _factors = []
    for _lbl, _row in row_for_label.items():
        if _is_non_currency_label(norm_api[_lbl]):
            continue
        _pv = ws.cell(_row, prev_col).value
        if isinstance(_pv, (int, float)) and _pv != 0 and raw_labels[_lbl] != 0:
            _f = infer_scale_factor(raw_labels[_lbl], _pv, tolerance_mult=3.0 if cumulative_flow else 1.0)
            if _f is not None:
                _factors.append(_f)
    default_factor = Counter(_factors).most_common(1)[0][0] if _factors else 1.0

    written = []
    for api_label, row in row_for_label.items():
        val = raw_labels[api_label]
        norm = norm_api[api_label]
        final_val = val
        if not _is_non_currency_label(norm):
            prev_val = ws.cell(row, prev_col).value
            prev_val = prev_val if isinstance(prev_val, (int, float)) else None
            if prev_val in (None, 0):
                target_cell = ws.cell(row, new_col, value=val / default_factor)
                copy_cell_style(ws.cell(row, prev_col), target_cell)
                written.append(api_label)
                continue
            if prev_val is not None and prev_val != 0 and val != 0 and (val > 0) != (prev_val > 0):
                continue  # тэмдэг зөрчилдөж байна — буруу мөр рүү тааруулсан байх магадлалтай
            factor = infer_scale_factor(val, prev_val, tolerance_mult=3.0 if cumulative_flow else 1.0)
            if factor is None:
                # Масштаб ЯМАР Ч байдлаар тохирохгүй байна — өмнөх утгатай
                # харьцуулахад хэт хэвийн бус тул буруу мөр рүү тааруулсан
                # байх магадлал өндөр. Худал тоо бичихээс илүү хоосон
                # орхих нь дээр (Telegram-аар "тохирол олдоогүй" гэж гарна).
                continue
            final_val = val / factor
        target_cell = ws.cell(row, new_col, value=final_val)
        copy_cell_style(ws.cell(row, prev_col), target_cell)
        written.append(api_label)

    all_sheet_rows = set(sheet_labels.values())
    still_empty_rows = all_sheet_rows - {row_for_label[l] for l in written}
    still_empty_labels = [k for k, v in sheet_labels.items() if v in still_empty_rows]
    return written, still_empty_labels


# ── 4. LibreOffice-оор бүх томьёог дахин тооцоолуулах (бие даасан) ─────
def find_soffice_binary() -> str | None:
    """Linux (soffice), macOS (LibreOffice.app), Windows (Program Files)
    аль ч дээр LibreOffice-ийг автоматаар олно (Gemini-ийн санаа)."""
    import shutil
    binary = shutil.which("soffice") or shutil.which("soffice.exe")
    if binary:
        return binary
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


_RECALC_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()
      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def recalc_workbook(path: str, timeout: int = 120):
    """
    LibreOffice headless-ийг нэг удаагийн макро script-ээр дуудаж бүх
    томьёог дахин тооцоолуулна (Дашборд, Харьцуулалт, CAMELS зэрэг энэ
    баланс/орлого хуудсуудаас шууд томьёогоор утга татдаг бусад бүх
    хуудас автоматаар шинэчлэгдэнэ). Ганц файл дотор бүрэн бие даасан —
    гадны скрипт (тухайлбал Claude-ийн орчны recalc.py) шаардахгүй.
    """
    import tempfile

    soffice_path = find_soffice_binary()
    if soffice_path is None:
        print("  ⚠ 'soffice' (LibreOffice) олдсонгүй.")
        print("    Debian/Ubuntu: sudo apt install libreoffice-calc")
        print("    macOS: brew install --cask libreoffice")
        print("    Windows: https://www.libreoffice.org/download/ -ээс суулгана уу")
        print("    Эсвэл файлыг гараар нээгээд Ctrl+Shift+F9 дараад хадгална уу.")
        return

    abs_path = str(Path(path).resolve())
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile_dir:
        profile_url = Path(profile_dir).as_uri()

        # 1. Профайл үүсгэх (эхний богино --terminate_after_init дуудлага)
        subprocess.run(
            [soffice_path, "--headless", "--terminate_after_init",
             f"-env:UserInstallation={profile_url}"],
            capture_output=True, timeout=30,
        )

        # 2. Макро суулгах
        macro_dir = Path(profile_dir) / "user" / "basic" / "Standard"
        macro_dir.mkdir(parents=True, exist_ok=True)
        (macro_dir / "Module1.xba").write_text(_RECALC_MACRO, encoding="utf-8")

        # 3. Макро ажиллуулж, дахин тооцоолж, хадгалах
        cmd = [
            soffice_path, "--headless", "--norestore",
            f"-env:UserInstallation={profile_url}",
            "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
            abs_path,
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            print(f"  ⚠ {timeout}с дотор дуусаагүй — дахин ажиллуулж үзнэ үү")
            return

        if result.returncode != 0:
            print(f"  ⚠ LibreOffice алдаатай гарлаа: {(result.stderr or '')[:300]}")
        else:
            print("  ✓ Бүх томьёо дахин тооцоологдож, файл хадгалагдлаа")


# ── 4b. Доод урсгалын (downstream) хуудсуудыг шинэчлэх ─────────────────
# "Snapshot" хуудас: банк тус бүрт ГАНЦ л "сүүлийн үеийн" багана заасан
# томьёотой — зүгээр л хуучин баганы үсгийг шинэ баганы үсгээр солино.
SNAPSHOT_SHEETS = [
    "Дашборд", "Харьцуулалт (9 банк)", "Активын чанар", "CAMELS дүгнэлт",
    "Graham Number", "Гүйцэх — ногдол ашиг", "Өмчлөл ба гүйцэтгэл",
]

# "Trend" хуудас: он-цагийн дараалалтай, улирал бүр ӨӨР багана — иймд
# хуучин баганыг дарж бичихгүй, ХАЖУУД нь шинэ багана НЭМНЭ. Блок бүрийн
# эхний мөр (энэ бол 5 MSE-д бүртгэлтэй банкны эхлэх мөр — Ариг/Богд/
# Транс зэрэг бусад 4 банк арай доор байрладаг тул энэ бот тэдэнд хамаарахгүй)
TREND_BLOCKS = {
    "9 банк зэрэгцүүлэлт": [(6, "bal", 22), (17, "bal", 8), (28, "bal", 27), (39, "bal", 40)],
    "Ашгийн зэрэгцүүлэлт":  [(6, "inc", 13), (18, "inc", 5), (29, "inc", 7)],
    "9 банк ашиг":          [(6, "inc", 13)],
    "Хүүгийн маржин":       [(6, "inc", 7), (18, "inc", None)],  # 18-р мөр 2 эшлэлтэй харьцаа
}
BANK_TREND_OFFSET = {"Төрийн банк": 0, "ХасБанк": 1, "ХХБ": 2, "Голомт": 3, "Хаан банк": 4}

# "Зохистой харьцаа" хуудас — БУСАД trend хуудсаас ялгаатай нь: томьёо
# ХУУЛБАРЛАХГҮЙ, харин Монголбанкны норматив харьцааг ШУУД УТГААР бичдэг
# (банкны улирлын тайлангийн "ЗОХИСТОЙ ХАРЬЦААНЫ ШАЛГУУР ҮЗҮҮЛЭЛТИЙН
# БИЕЛЭЛТ" хэсэгтэй яг ижил 5 харьцаа). Блок гарчгийг шошготой уялдуулна.
RATIO_SHEET = "Зохистой харьцаа"
RATIO_BLOCKS = [
    (5, "нэгдүгээр зэрэглэлийн өөрийн хөрөнгө болон эрсдэлээр жигнэсэн активын зохистой харьцаа"),
    (17, "өөрийн хөрөнгө болон эрсдэлээр жигнэсэн активын харьцаа"),
    (29, "төлбөр түргэн гүйцэтгэх чадварын үзүүлэлт"),
    (41, "гадаад валютын нээлттэй позици болон өөрийн хөрөнгийн харьцаа (нэгж валютаар)"),
    (53, "гадаад валютын нээлттэй позици болон өөрийн хөрөнгийн харьцаа (нийт)"),
]


def normalize_ratio_label(s: str) -> str:
    s = re.sub(r"\[.*?\]", "", str(s))  # "[Байвал зохих: > 9%]" гэсэн хаалтыг хас
    return re.sub(r"\s+", " ", s).strip().lower()


def update_ratio_sheet(wb, bank: str, date_label: str, raw_labels: dict[str, float | None]) -> int:
    """Зохистой харьцааны 5 блокт тухайн банкны мөрөнд шинэ (эсвэл байгаа)
    багана дээр утга бичнэ. API-аас ирсэн утга % хэлбэртэй (жиш. 16.0)
    бол 0-1 хоорондох бутархай (0.16) рүү хөрвүүлнэ; аль хэдийн бутархай
    (жиш. 0.16) бол хэвээр үлдээнэ."""
    if RATIO_SHEET not in wb.sheetnames:
        return 0
    offset = BANK_TREND_OFFSET.get(bank)
    if offset is None:
        return 0
    ws = wb[RATIO_SHEET]
    trend_label = trend_date_label(date_label)

    # Зорилтот баганыг НЭГ УДАА тодорхойл (аль ч блокийн толгой мөрөөс
    # хайж болно — таарвал БҮХ 5 блок ЯГ ТЭР НЭГ баганад бичигдэнэ)
    target_col = None
    for c in range(2, ws.max_column + 1):
        if any(str(ws.cell(hr, c).value).strip() == trend_label for hr, _ in RATIO_BLOCKS):
            target_col = c
            break
    if target_col is None:
        target_col = ws.max_column + 1

    norm_raw = {normalize_ratio_label(k): v for k, v in raw_labels.items() if v is not None}

    changed = 0
    for header_row, block_label_norm in RATIO_BLOCKS:
        val = norm_raw.get(block_label_norm)
        if val is None:
            for k, v in norm_raw.items():
                if block_label_norm in k or k in block_label_norm:
                    val = v
                    break
        if val is None:
            # Бичлэгийн жижиг зөрүүтэй (жиш. "позиц" vs "позици", "болон
            # эрсдэлээр" vs "эрсдлээр") тохиолдолд үгийн багц давхцлаар
            # олно — бодит XacBank файл дээр яг ийм ялгаа олдсон.
            block_words = set(block_label_norm.split())
            best_v, best_score = None, (0.0, 0)
            for k, v in norm_raw.items():
                k_words = set(k.split())
                if not k_words:
                    continue
                overlap = len(block_words & k_words) / len(block_words)
                # Ижил overlap-тай бол ИЛҮҮ ҮГГҮЙ (богино) шошгыг илүүд үзнэ: Хаан
                # банкны "Өөрийн хөрөнгө, эрсдэлээр жигнэсэн активын харьцаа" (CAR 18.9%)
                # нь "Нэгдүгээр зэрэглэлийн … харьцаа" (Tier1 18.1%)-тай ижил overlap
                # өгч, Tier1-ийн утга CAR-д бичигдэж байсныг Q3 дээр олсон.
                score = (overlap, -len(k_words))
                if overlap >= 0.6 and score > best_score:
                    best_v, best_score = v, score
            val = best_v
        # Толгойг ЗААВАЛ бич (утга олдоогүй ч бусад 4 блоктой зэрэгцэхийн тулд)
        ws.cell(header_row, target_col, value=trend_label)
        if val is None:
            continue  # энэ харьцаа эх сурвалжид алга — хоосон орхино, дараа нь мэдэгдэнэ

        decimal_val = val / 100.0 if abs(val) > 1.5 else val  # "16.0" (%) -> 0.16

        r = header_row + 1 + offset
        cell = ws.cell(r, target_col, value=decimal_val)
        copy_cell_style(ws.cell(r, target_col - 1), cell)
        changed += 1
    return changed


def fix_khasbank_cumulative_to_standalone(ws, header_row: int, new_col: int, date_label: str = "") -> int:
    """
    ХасБанкны орлого энэ workbook-д УЛИРЛЫН ДАН ДҮНгээр хадгалагддаг, харин
    МХБ-ийн тайлан оны эхнээс ХУРИМТЛАЛААР (Q2=H1, Q3=9 сар, Q4=жил) ирдэг.
    Иймд дан дүн = шинэ хуримтлал − ТУХАЙН ОНЫ өмнөх БҮХ улирлын дан дүнгийн
    нийлбэр (Q1: юу ч хасахгүй; Q2: Q1; Q3: Q1+Q2; Q4: Q1+Q2+Q3).
    Өмнөх хувилбар зөвхөн өмнөх баганыг хасдаг байсан нь Q3/Q4-д буруу.
    """
    year = str(date_label)[:4]
    prior_cols = []
    for c in range(2, new_col):
        h = str(ws.cell(header_row, c).value or "").strip()
        if year and h.startswith(year):
            prior_cols.append(c)
    if not prior_cols:
        return 0  # Q1 — хуримтлал нь дан дүнтэй тэнцүү
    fixed = 0
    for r in range(header_row + 1, ws.max_row + 1):
        new_val = ws.cell(r, new_col).value
        if not isinstance(new_val, (int, float)):
            continue
        prior = [ws.cell(r, c).value for c in prior_cols]
        if not all(isinstance(p, (int, float)) for p in prior):
            continue
        ws.cell(r, new_col, value=new_val - sum(prior))
        fixed += 1
    return fixed


ANNUALIZE_SHEETS = ["Дашборд", "Харьцуулалт (9 банк)", "CAMELS дүгнэлт", "Гүйцэх — ногдол ашиг"]
ANNUAL_FACTOR = {1: "4", 2: "2", 3: "(4/3)", 4: "1"}
_ANN_RE = re.compile(r"\*\s*(?:\(4/3\)|[124])(?![\d.])")


def update_annualization(wb, bank: str, quarter: int) -> int:
    """
    ROE/ROA нь тайлант үеийн (хуримтлал) ашгийг жил болгон хөрвүүлдэг: H1 ×2
    гэх мэт. Томьёонд "*2" шууд бичигдсэн тул Q3 (9 сар) орсон ч ×2 хэвээр
    үлдэж ROE-г хэт өсгөдөг (Төрийн банк 27.4% vs зөв 18.2%) гэдгийг бодит
    Q3 файл дээр олсон. Банкны томьёонд (нэр заасан эсвэл баганаар) өнөөгийн
    улирлын коэффициентийг (Q1×4, Q2×2, Q3×4/3, Q4×1) тавина.
    """
    factor = ANNUAL_FACTOR[quarter]
    changed = 0
    for sheet in ANNUALIZE_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=") and _ANN_RE.search(v)):
                    continue
                names = [b for b in BANK_TREND_OFFSET if f"'{b} " in v]
                if names:
                    owner = names[0]
                elif sheet == "Харьцуулалт (9 банк)" and 2 <= cell.column <= 6:
                    owner = list(BANK_TREND_OFFSET)[cell.column - 2]
                else:
                    continue
                if owner != bank:
                    continue
                new_v = _ANN_RE.sub(f"*{factor}", v)
                if new_v != v:
                    cell.value = new_v
                    changed += 1
    return changed


YTD_SHEETS = ["Дашборд", "Харьцуулалт (9 банк)", "CAMELS дүгнэлт", "Graham Number", "Гүйцэх — ногдол ашиг"]
_KHAS_YTD_RE = re.compile(r"'ХасБанк орлого'![A-Z]+13(?:\+'ХасБанк орлого'![A-Z]+13)*")


def rebuild_khasbank_ytd(wb, date_label: str) -> int:
    """
    ХасБанкны орлого УЛИРЛЫН ДАН ДҮНгээр хадгалагддаг тул жилийн (YTD) ашиг нь
    оны Q1+Q2+... нийлбэр: "('ХасБанк орлого'!K13+'ХасБанк орлого'!L13)".
    Ерөнхий repoint нь зөвхөн сүүлийн баганы үсгийг солидог тул Q3-д Q1+Q3
    (Q2 алгасагдсан), эсвэл Q2+Q2 болж ХасБанкны ROE/ашиг буруу гардаг байсныг
    олсон. Одоо тухайн оны БҮХ улирлын багануудаар (Q1..Qn) дахин барина.
    """
    ws_inc = wb["ХасБанк орлого"]
    hdr = HEADER_ROW["ХасБанк"]["inc"]
    year = str(date_label)[:4]
    new_col = ws_inc.max_column
    letters = [col_letter(c) for c in range(2, new_col + 1)
               if str(ws_inc.cell(hdr, c).value or "").strip().startswith(year)]
    if not letters:
        return 0
    chain = "+".join(f"'ХасБанк орлого'!{L}13" for L in letters)
    changed = 0
    for sheet in YTD_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        for row in wb[sheet].iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("=") and _KHAS_YTD_RE.search(v):
                    new_v = _KHAS_YTD_RE.sub(chain, v)
                    if new_v != v:
                        cell.value = new_v
                        changed += 1
    return changed


PERIOD_LABEL_CELLS = {
    "Тайлбар": ["B5"],
    "Дашборд": ["A2", "A4", "A8", "A14"],
    "Харьцуулалт (9 банк)": ["A2", "A4", "A11", "A12", "A18"],
    "Активын чанар": ["A2", "A4"],
    "CAMELS дүгнэлт": ["A8"],
    "Гүйцэх — ногдол ашиг": ["A4"],
}
_PERIOD_NAME = {1: "3 сар", 2: "H1", 3: "9 сар", 4: "FY"}
NON_MSE_BANKS = "Капитрон, Ариг, Богд, Транс банк"


def update_period_labels(wb, old_date: str, new_date: str, all_five_updated: bool = True) -> int:
    """
    Дашборд, Харьцуулалт, Активын чанар, CAMELS, Гүйцэх, Тайлбар хуудсуудын
    ГАРЧИГТ гараар бичигдсэн огноог ("2026/06/30", "2026 H1", "FY2025") шинэчилнэ.
    Бот үүнийг өмнө нь огт хийдэггүй байсан тул Q3 workbook-д хуучин огноо үлдсэн.
    МХБ-ийн бус 4 банк (бот хамаарахгүй) хоцорч байдаг тул "Бүх 9 банк" гэсэн
    мэдэгдлийг "МХБ-ийн 5 банк <шинэ>; бусад 4 банк <хуучин>" болгоно. Зөвхөн
    5 банк бүгд шинэчлэгдсэн үед (аль нэг нь хоцорсон бол гараар) ажиллана.
    """
    if not all_five_updated:
        print("  ⚠ 5 банк бүгд шинэчлэгдээгүй тул гарчгийн огноог гараар шинэчилнэ үү")
        return 0
    old_d, new_d = old_date.replace("*", ""), new_date.replace("*", "")
    year, month = new_d[:4], int(new_d[5:7])
    quarter = month // 3
    pname = _PERIOD_NAME[quarter]
    period = f"{year} {pname}" if quarter < 4 else f"FY{year}"
    changed = 0
    for sheet, cells in PERIOD_LABEL_CELLS.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for coord in cells:
            v = ws[coord].value
            if not isinstance(v, str) or v.startswith("="):
                continue
            s = v.replace(old_d, new_d)
            s = re.sub(r"Цэвэр ашиг (?:FY\d{4}|\d{4} (?:3 сар|H1|9 сар))(?: \(YTD\))?",
                       f"Цэвэр ашиг {period} (YTD)", s)
            s = re.sub(r"Цэвэр ашиг \((?:\d{4} (?:3 сар|H1|9 сар)|FY\d{4})\)", f"Цэвэр ашиг ({period})", s)
            s = re.sub(r"Бүх 9 банк( —)? ?(" + re.escape(new_d) + r")",
                       lambda m: f"МХБ-ийн 5 банк{m.group(1) or ''} {m.group(2)}; {NON_MSE_BANKS} — {old_d}", s)
            if s != v:
                ws[coord].value = s
                changed += 1
    return changed


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def repoint_snapshot_formulas(wb, bank: str, old_col_bal: int, new_col_bal: int,
                               old_col_inc: int, new_col_inc: int) -> int:
    """SNAPSHOT_SHEETS доtorh, тухайн банкны баланс/орлого хуудас руу
    заасан бүх томьёог хуучин баганаас шинэ багана руу дахин чиглүүлнэ."""
    old_b, new_b = col_letter(old_col_bal), col_letter(new_col_bal)
    old_i, new_i = col_letter(old_col_inc), col_letter(new_col_inc)
    pat_bal = re.compile(rf"('{re.escape(bank)} баланс'!){old_b}(\d+)")
    pat_inc = re.compile(rf"('{re.escape(bank)} орлого'!){old_i}(\d+)")
    changed = 0
    for sheet in SNAPSHOT_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    new_val = pat_bal.sub(rf"\g<1>{new_b}\2", cell.value)
                    new_val = pat_inc.sub(rf"\g<1>{new_i}\2", new_val)
                    if new_val != cell.value:
                        cell.value = new_val
                        changed += 1
    return changed


def trend_date_label(date_label: str) -> str:
    """Эх sheet-ийн 'YYYY/MM/DD[*]' огноог trend хуудасны 'YYYY/MM' форматад хөрвүүлнэ."""
    return date_label.replace("*", "")[:7]


def add_trend_columns(wb, bank: str, date_label: str,
                       old_col_bal: int, new_col_bal: int,
                       old_col_inc: int, new_col_inc: int) -> int:
    """TREND_BLOCKS-д заасан хуудсуудад тухайн банкны мөрөнд шинэ багана
    нэмнэ. ЧУХАЛ: баганы үсгийг trend-хуудасны ӨМНӨХ баганаас "+1" гэж
    ТААМАГЛАХГҮЙ — учир нь зарим банк (жиш. Хаан банк) эх баланс/орлого
    хуудсандаа түүхэн шалтгаанаар илүү баганатай (dutuu улирал нөхсөн)
    тул баганы индекс бусад банктай зэрэгцэхгүй байж болно. Оронд нь
    main()-с ирсэн БОДИТ old_col/new_col-ыг ашиглаж, тэдгээрийн ЯГ ТУХАЙН
    баганы үсгийг л томьёон дотор орлуулна.
    """
    offset = BANK_TREND_OFFSET.get(bank)
    if offset is None:
        return 0
    trend_label = trend_date_label(date_label)
    old_b, new_b = col_letter(old_col_bal), col_letter(new_col_bal)
    old_i, new_i = col_letter(old_col_inc), col_letter(new_col_inc)
    changed = 0
    for sheet, blocks in TREND_BLOCKS.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        header_row = blocks[0][0] - 1  # "Банк / Улирал" мөр нь эхний банкны мөрөөс яг 1 дээш

        # Энэ огнооны багана аль хэдийн байгаа эсэхийг header мөрөөс шалгах
        target_col = None
        for c in range(2, ws.max_column + 1):
            if str(ws.cell(header_row, c).value).strip() == trend_label:
                target_col = c
                break
        prev_col = target_col - 1 if target_col else ws.max_column
        if target_col is None:
            target_col = ws.max_column + 1

        for block_row, kind, _src_row in blocks:
            r = block_row + offset
            old_formula = ws.cell(r, prev_col).value
            if not (isinstance(old_formula, str) and old_formula.startswith("=")):
                continue
            old_letter, new_letter = (old_b, new_b) if kind == "bal" else (old_i, new_i)
            if not re.search(rf"!{old_letter}\d+", old_formula):
                continue  # өмнөх баганын томьёо өөр эх багана заасан — дарж бичихгүй
            new_formula = re.sub(rf"!{old_letter}(\d+)", rf"!{new_letter}\1", old_formula)
            new_cell = ws.cell(r, target_col, value=new_formula)
            copy_cell_style(ws.cell(r, prev_col), new_cell)
            changed += 1

        if header_row >= 1:
            h_cell = ws.cell(header_row, target_col, value=trend_label)
            copy_cell_style(ws.cell(header_row, prev_col), h_cell)
    return changed


# ── 5. Telegram мэдэгдэл ────────────────────────────────────────────────
def send_telegram_message(text: str):
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("  ℹ TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID тохируулаагүй тул Telegram алгасав")
        print("  --- Мэдэгдлийн агуулга ---")
        print(text)
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        resp = requests.post(url, data={"chat_id": TELEGRAM_CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=15)
        resp.raise_for_status()
        print("  ✓ Telegram мэдэгдэл илгээгдлээ")
    except Exception as e:
        print(f"  ⚠ Telegram илгээхэд алдаа гарлаа: {e}")


# ── 6. Үндсэн урсгал ─────────────────────────────────────────────────────
def resolve_excel_path() -> str | None:
    """
    Jupyter/Colab орчинд шууд (jupyter-ийн % эсвэл import-оор) ажиллуулбал
    sys.argv[1] "-f /root/.local/share/jupyter/runtime/kernel-xxx.json" гэсэн
    Jupyter-ийн ӨӨРИЙН аргумент байдаг — таны ЭНЭ ГАРСАН "-f олдсонгүй" алдаа
    яг үүнээс болсон. Иймд эхлээд argv-аас бодит .xlsx зам хайж, олдохгүй
    бол одоогийн хавтас доtorh цорын ганц .xlsx файлыг автоматаар олж авна.
    """
    for a in sys.argv[1:]:
        if a.lower().endswith(".xlsx") and Path(a).exists():
            return a
    xlsx_files = [p for p in Path(".").glob("*.xlsx") if not p.name.startswith("~$")]
    if len(xlsx_files) == 1:
        print(f"ℹ Excel файлын зам аргументаар өгөгдөөгүй тул автоматаар олдсон: {xlsx_files[0]}")
        return str(xlsx_files[0])
    if len(xlsx_files) > 1:
        print("❌ Одоогийн хавтаст хэд хэдэн .xlsx файл байна, аль нэгийг нь тодорхой зааж өгнө үү:")
        for p in xlsx_files:
            print(f"   - {p}")
    return None


def main():
    check_only = "--check" in sys.argv
    excel_path = resolve_excel_path()

    if excel_path is None:
        print("Ашиглах:")
        print("  Терминалаас:  python3 bank_financials_bot.py <excel_file.xlsx> [--check]")
        print("  Colab/Jupyter-ээс (АНХААР: ! тэмдэг заавал хэрэгтэй — commands ажиллуулна,")
        print("  Python код гэж биш):")
        print('    !python3 bank_financials_bot.py "Есөн_банк_шинжилгээ2026-06-30.xlsx"')
        print("  Эсвэл: файлаа скрипттэй ЯГ НЭГ хавтаст байрлуулаад аргументгүйгээр ажиллуулж болно")
        print("  (тэр үед цорын ганц .xlsx файлыг автоматаар олно).")
        sys.exit(1)

    if not Path(excel_path).exists():
        print(f"❌ {excel_path} олдсонгүй")
        sys.exit(1)

    wb = load_workbook(excel_path)
    telegram_lines = [f"📊 <b>Банкны санхүүгийн тайлангийн шалгалт</b> — {datetime.now().strftime('%Y-%m-%d %H:%M')}"]
    any_updated = False
    updated_banks: dict = {}

    for bank, info in BANKS.items():
        print(f"\n{'='*60}\n🔍 {bank} ({info['ticker']}, id={info['id']})")

        # 1. Аль улирлын тайлан хамгийн сүүлд гарсныг тодорхойлно (зөвхөн
        #    он/улирал, ӨГӨГДӨЛ АВАХГҮЙ)
        site_raw = None
        if bank == "Голомт" and "--golomt-site" in sys.argv:
            # МХБ-д Голомтын шинэ улирлын файл гараагүй үед банкны өөрийн IR
            # хуудсыг (master-ийн Голомт sheet-ийг бүтээсэн яг тэр загвар) эх
            # сурвалж болгоно. Хамгийн сүүлийн улирлыг тэр хуудаснаас авна.
            try:
                site = fetch_golomt_site()
            except Exception as e:
                site = None
                print(f"  ⚠ Голомтын сайтаас татахад алдаа: {e}")
            if not site:
                telegram_lines.append(f"⚠️ {bank}: сайтаас тайлан олдсонгүй")
                continue
            api_date, site_raw = site
            print(f"  Голомтын сайт дээрх сүүлийн тайлан: {api_date} ({len(site_raw)} шошго)")
        else:
            try:
                quarter_info = fetch_quarter_info(info["id"])
            except Exception as e:
                print(f"  ⚠ Улирал тодорхойлоход алдаа гарлаа: {e}")
                telegram_lines.append(f"⚠️ {bank}: улирал тодорхойлоход алдаа ({e})")
                continue

            if not quarter_info:
                print("  ⚠ Улирлын мэдээлэл олдсонгүй (хуудасны бүтэц өөрчлөгдсөн байж болзошгүй)")
                telegram_lines.append(f"⚠️ {bank}: улирал тодорхойлогдсонгүй")
                continue

            api_date = quarter_end_date(quarter_info["year"], quarter_info["quarter"])
            print(f"  MSE дээрх сүүлийн тайлан: {quarter_info['heading']} -> {api_date}")

        # 2. Excel дэх сүүлийн огноо
        ws_bal = wb[f"{bank} баланс"]
        ws_inc = wb[f"{bank} орлого"]
        hdr_bal = HEADER_ROW[bank]["bal"]
        hdr_inc = HEADER_ROW[bank]["inc"]
        excel_date_bal, last_col_bal = get_excel_last_date(ws_bal, hdr_bal)
        excel_date_inc, last_col_inc = get_excel_last_date(ws_inc, hdr_inc)
        print(f"  Excel дэх сүүлийн огноо: {excel_date_bal}")

        # 3. Харьцуулах — шинэ бол л үргэлжлүүлнэ
        if api_date.replace("*", "") <= excel_date_bal.replace("*", ""):
            print(f"  ✓ Аль хэдийн сүүлийн үеийн мэдээлэлтэй байна, алгасав")
            telegram_lines.append(f"✓ {bank}: {excel_date_bal} (шинэчлэл шаардлагагүй)")
            continue

        print(f"  🌟 ШИНЭ тайлан олдлоо: {excel_date_bal} -> {api_date}")
        any_updated = True

        # 1b. Албан ёсны .xls/.xlsx тайлан файлыг татна — ЦОРЫН ГАНЦ эх
        #     сурвалж (HTML хуудаснаас ямар ч утга авахгүй)
        if site_raw is not None:
            report = {"raw_labels": site_raw, "source": GOLOMT_SITE_URL}
        else:
            try:
                report = fetch_full_report(info["id"])
            except Exception as e:
                report = None
                print(f"  ⚠ Тайлан файл татахад алдаа гарлаа: {e}")

        if not report or not report.get("raw_labels"):
            print(f"  ❌ Албан ёсны тайлан файл олдсонгүй эсвэл задлагдсангүй — "
                  f"энэ банкийг алгасав (утгагүй/дутуу өгөгдөл бичихгүй)")
            telegram_lines.append(f"❌ {bank}: {excel_date_bal} → {api_date} — "
                                   f"тайлан файл олдсонгүй, шинэчлэгдсэнгүй")
            continue

        raw_labels = report["raw_labels"]
        print(f"  ✓ Тайлан файлаас {len(raw_labels)} шошго задлагдлаа (эх: {report['source']})")

        if check_only:
            telegram_lines.append(f"🆕 {bank}: {excel_date_bal} -> {api_date} (--check горим, бичихгүй)")
            continue

        # 4. Бичих — тайлан файлаас ирсэн БҮХ талбарыг Excel-ийн бодит
        #    мөрийн шошготой тааруулж, боломжтой болгоныг нь бичнэ
        new_col_bal = last_col_bal + 1
        new_col_inc = last_col_inc + 1
        written_bal, empty_bal = write_new_quarter_full(ws_bal, hdr_bal, new_col_bal, api_date, raw_labels)
        written_inc, empty_inc = write_new_quarter_full(
            ws_inc, hdr_inc, new_col_inc, api_date, raw_labels,
            cumulative_flow=(bank == "ХасБанк"))

        if bank == "ХасБанк":
            n_fixed = fix_khasbank_cumulative_to_standalone(ws_inc, hdr_inc, new_col_inc, api_date)
            print(f"  ✓ ХасБанк орлого: {n_fixed} мөрийг H1-хуримтлалаас Q2-дан дүн рүү хөрвүүлэв")

        print(f"  ✓ Баланс: {len(written_bal)} мөр бичигдлээ | Орлого: {len(written_inc)} мөр бичигдлээ")
        if empty_bal or empty_inc:
            print(f"  ⚠ Тохирол олдоогүй, хоосон үлдсэн мөрүүд: "
                  f"баланс={empty_bal[:5]}{'...' if len(empty_bal)>5 else ''}, "
                  f"орлого={empty_inc[:5]}{'...' if len(empty_inc)>5 else ''}")

        # 4b. Доод урсгалын хуудсуудыг шинэчлэх
        n_repointed = repoint_snapshot_formulas(wb, bank, last_col_bal, new_col_bal, last_col_inc, new_col_inc)
        if bank == "ХасБанк":
            n_ytd = rebuild_khasbank_ytd(wb, api_date)
            print(f"  ✓ ХасБанк YTD ашгийн нийлбэр: {n_ytd} томьёо дахин баригдлаа")
        updated_banks[bank] = (excel_date_bal, api_date)
        n_ann = update_annualization(wb, bank, int(api_date[5:7]) // 3)
        print(f"  ✓ Жилчлэлийн коэффициент: {n_ann} томьёо шинэчлэгдлээ")
        n_trend = add_trend_columns(wb, bank, api_date, last_col_bal, new_col_bal, last_col_inc, new_col_inc)
        n_ratio = update_ratio_sheet(wb, bank, api_date, raw_labels) if UPDATE_RATIO_SHEET else 0
        print(f"  ✓ Дашборд/Харьцуулалт г.м. {n_repointed} нүд дахин чиглүүлэгдлээ; "
              f"trend хуудсанд {n_trend} шинэ багана нэмэгдлээ; "
              + (f"Зохистой харьцаанд {n_ratio}/5 харьцаа бичигдлээ" if UPDATE_RATIO_SHEET
                 else "Зохистой харьцаа хөндөгдөөгүй"))

        empty_summary = ""
        if empty_bal or empty_inc:
            all_empty = (empty_bal + empty_inc)[:8]
            empty_summary = f"\n   ⚠️ Тохирол олдоогүй ({len(empty_bal)+len(empty_inc)} мөр): {', '.join(all_empty)}{'...' if len(empty_bal)+len(empty_inc) > 8 else ''} — эдгээрийг гараар шалгана уу."
        telegram_lines.append(
            f"🆕 <b>{bank}</b>: {excel_date_bal} → <b>{api_date}</b>\n"
            f"   Баланс: {len(written_bal)} мөр, Орлого: {len(written_inc)} мөр амжилттай бичигдлээ."
            f"{empty_summary}"
        )

    if check_only:
        print("\n--check горим тул хадгалаагүй.")
        send_telegram_message("\n\n".join(telegram_lines))
        return

    if any_updated:
        olds = {o for o, _ in updated_banks.values()}
        news = {n for _, n in updated_banks.values()}
        if len(olds) == 1 and len(news) == 1:
            n_lbl = update_period_labels(wb, next(iter(olds)), next(iter(news)),
                                         all_five_updated=len(updated_banks) == len(BANKS))
            print(f"  ✓ Гарчгийн огноо: {n_lbl} нүд шинэчлэгдлээ")
        wb.save(excel_path)
        print(f"\n✓ {excel_path} хадгалагдлаа")
        print("\n🔄 Дахин тооцоолж байна (Дашборд, Харьцуулалт г.м. бусад хуудсууд шинэчлэгдэнэ)...")
        recalc_workbook(excel_path)
    else:
        print("\nℹ Ямар ч банкинд шинэ тайлан олдсонгүй — файл өөрчлөгдөөгүй")

    send_telegram_message("\n\n".join(telegram_lines))


if __name__ == "__main__":
    main()
