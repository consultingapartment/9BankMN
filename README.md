# МХБ 5 Банкны Dashboard (автомат шинэчлэгддэг)

МХБ-д бүртгэлтэй 5 банкны (Төрийн банк, ХасБанк, ХХБ, Голомт, Хаан банк)
техникийн ба суурь шинжилгээний интерактив HTML dashboard. GitHub Actions
өдөр бүр автоматаар `old.mse.mn`-ээс шинэ ханш татаж, сайтыг дахин
шинэчилж, GitHub Pages дээр нийтэлдэг.

## Хэрхэн ажилладаг вэ (архитектур)

```
┌──────────────────┐   cron (Даваа-Баасан 15:30 УБ цаг)
│ GitHub Actions    │──────────────────┐
└──────────────────┘                   ▼
        │                 scripts/fetch_prices.py
        │                 (old.mse.mn → data/price_data.json)
        │                              │
        │                              ▼
        │                 scripts/build_site.py
        │                 (template + data → docs/index.html)
        │                              │
        └──────── git commit + push ───┘
                                        │
                                        ▼
                              GitHub Pages (docs/)
                              → https://<user>.github.io/<repo>/
```

Telegram, alert зэрэг нэмэлт мессеж илгээхгүй — зөвхөн **статик сайтыг
автоматаар шинэчилж байлгах** энгийн, найдвартай урсгал.

## Анхны тохиргоо

1. Энэ хавтсыг шинэ GitHub repo болгож push хий (**Binance Square bot-той
   ХОЛИХГҮЙ, тусдаа repo байх ёстой**).
2. GitHub repo → Settings → Pages → Source: `Deploy from a branch`,
   Branch: `main` / `docs` хавтас сонго.
3. Settings → Actions → General → Workflow permissions →
   **"Read and write permissions"** идэвхжүүл (Actions-аас commit хийхэд
   шаардлагатай).
4. Анх удаад гараар ажиллуулж шалга: Actions таб → "Update MSE bank
   dashboard" → "Run workflow".

## Файлын бүтэц

- `scripts/fetch_prices.py` — old.mse.mn-ээс 5 банкны OHLCV татаж,
  `data/price_data.json`-той нэгтгэнэ. Компанийн ID = ҮЦ код
  (562=GLMT, 563=KHAN, 564=SBM, 567=TDB, 568=XAC).
- `scripts/build_site.py` — `template/dashboard_template.html`-д
  `data/*.json`-ыг шингээж `docs/index.html` болгоно.
- `scripts/indicators.py` — MA/RSI/MACD/S-R/SMC/Wyckoff-ийн Python
  логик (одоогоор dashboard дотор JavaScript-аар давхар тооцоологддог
  тул ЗААВАЛ хэрэггүй; ирээдүйд Python талд алерт/шинжилгээ нэмэх
  бол бэлэн байгаа болгож үлдээв).
- `docs/index.html` — GitHub Pages-ээр нийтлэгдэх эцсийн файл
  (автоматаар үүсдэг тул ГАРААР бүү засвар хий — дараагийн шинэчлэлтэд
  дарагдана).
- `docs/Есөн_банк_шинжилгээ.xlsx` — 9 банкны бүрэн санхүүгийн
  шинжилгээ (Excel), сайтын толгойгоос татаж авах боломжтой.

## Гараар шинэчлэх (тест хийхэд)

```bash
pip install -r requirements.txt
python3 scripts/fetch_prices.py      # эсвэл --check зөвхөн одоогийн байдал харах
python3 scripts/build_site.py
```

## Мэдэгдэж буй хязгаарлалт

- `old.mse.mn`-ийн хуудасны бүтэц өөрчлөгдвөл `fetch_prices.py`-ийн
  regex ажиллахгүй болно — Actions run амжилтгүй болсноор мэдэгдэнэ.
- Амралтын өдөр/арилжаагүй өдөр шинэ мэдээлэл ирэхгүй тул
  `git diff --cached --quiet` шалгалт ажиллаж дэмий commit хийхгүй.

---

# Улирлын тайлан → шинжилгээ (автомат)

`.github/workflows/quarterly.yml` өдөр бүр 09:17 (УБ) ажиллана:

1. **Илрүүлэх** — open.mse.mn дээр 5 банкны сүүлийн тайлан workbook-оос шинэ эсэхийг шалгана.
   Голомт МХБ-д хоцорвол golomtbank.com IR хуудаснаас автоматаар авна.
2. **Татах** — тухайн улирлын блок доторх БҮХ .xls/.xlsx файлыг татаж `reports/<YYYY>Q<n>/`-д архивлана
   (sha256-тай, задалсан шошгууд `*.labels.json`). Хэд хэдэн файл байвал огноо таарсан, хамгийн бүрэн
   задлагдсаныг сонгоно. Файлд өөр улирлын огноо байвал **бичихгүй**.
3. **Бичих** — `bank_financials_bot.py`-ийн функцээр шинэ багана; Дашборд г.м. томьёог шинэ багана руу
   чиглүүлж, жилчлэл (Q1×4, H1×2, 9 сар×4/3, жил×1), ХасБанкны дан дүн, trend багана, гарчгийн огноо.
4. **Чанарын хаалт** — банк бүрт: Хөрөнгө = Өр + ӨХ, гол мөрүүд бичигдсэн, мөрийн ≥60% бичигдсэн,
   нийт хөрөнгө улирлаар −30%…+40% дотор. Унавал тухайн банкийг **буцаана**, Telegram-д шалтгааныг бичнэ.
5. **Recalc** — LibreOffice. Томьёоны алдаа өсвөл master-ийг солихгүй, `*_CANDIDATE.xlsx` үлдээнэ.
   Хуучин хувилбар `archive/`-д.
6. **Шинжилгээ** — `analysis/<YYYY>Q<n>/`: Word тайлан, `metrics.json/csv`, график; `data/fundamentals.json`
   (TTM EPS, BVPS) шинэчлэгдэж dashboard дахин бүтээгдэнэ. P/E, P/B-г dashboard өдөр бүрийн ханшаар тооцно.
7. **Telegram** — ЗӨВХӨН шинжилгээний Word файл (текст, график, тайлбаргүй). Тайлан орсон төлөв,
   алдаа зэрэг нь Word-ийн сүүлийн "Өгөгдлийн төлөв ба эх сурвалж" бүлэгт байна; шинжилгээ гараагүй
   өдөр (шинэ тайлангүй, эсвэл бүх банк шалгалтад унасан) юу ч илгээхгүй — Actions-ийн лог-д үлдэнэ.
   Тест: `python3 bank_pipeline/test_telegram_offline.py analysis/2026Q3`

Банкууд өөр өөр өдөр тайлангаа гаргасан ч болно — шинэ гарсныг нь л нэмнэ. 5 банк бүгд ижил огноонд
хүрэхэд гарчгийн огноо автоматаар шинэчлэгдэнэ. **Q4 аудитлагдсан тайлан** ижил улирлын файлыг солиход
(sha өөрчлөгдвөл) тухайн баганыг дахин бичнэ.

## Тохиргоо (нэг удаа)

- Settings → Secrets and variables → Actions: `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`
- Actions → "Quarterly bank reports" → **Run workflow** (`force_analysis` сонгож анх удаа туршина)

## Гараар / локал ажиллуулах

```bash
pip install -r requirements.txt      # + LibreOffice (recalc-д)
python3 bank_pipeline/quarterly_pipeline.py --workbook docs/Есөн_банк_шинжилгээ.xlsx --detect-only
python3 bank_pipeline/quarterly_pipeline.py --workbook docs/Есөн_банк_шинжилгээ.xlsx \
        --prices data/price_data.json --fundamentals data/fundamentals.json
# МХБ ажиллахгүй бол: татсан файлаа inbox/-д (нэр 564_*, SBM_*, эсвэл банкны нэрээр эхэлнэ)
python3 bank_pipeline/quarterly_pipeline.py --workbook docs/Есөн_банк_шинжилгээ.xlsx --inbox inbox --period 2026Q4
# Зөвхөн шинжилгээ
python3 bank_pipeline/quarterly_pipeline.py --workbook docs/Есөн_банк_шинжилгээ.xlsx --analysis-only \
        --prices data/price_data.json --fundamentals data/fundamentals.json
# Сүлжээгүй бүрэн тест (Q3 regression, Q4 симуляци, restatement, хаалт, огноо)
python3 bank_pipeline/test_pipeline_offline.py <Q2_master.xlsx> <Q3_workbook.xlsx>
```

## Капитрон банк (capitronbank.mn)

`bank_pipeline/capitron_source.py` нь банкны сайтын WordPress API-аас (ангилал 43, 44) тайлангуудыг
уншина. Гарчиг, агуулга эсвэл файлын нэрээс улирлыг таньж ("III улирал", "2-р улирал",
"2026.09.30", "хагас жил", "аудитлагдсан" → Q4), .pdf/.xlsx файлыг нийтлэлийн аль ч хэсгээс олно
(content, `_embedded`, acf, media). PDF-ийг текстээр нь задална (скан PDF-ийг алгасаж мэдэгдэнэ).
Стандарт маягтын мөрүүдийг Капитрон sheet-ийн өвөрмөц мөрүүдэд нэгтгэнэ (Харилцах + Хадгаламж →
"Харилцагчдын харилцах, хадгаламж" гэх мэт). Pipeline үүнийг МХБ-ийн 5 банктай хамт автоматаар
ажиллуулна: trend хуудас, CAMELS/Дашборд томьёо, жилчлэл бүгд шинэчлэгдэнэ.

```bash
python3 bank_pipeline/capitron_source.py --list          # олдсон тайлан, улирал, файлууд
python3 bank_pipeline/capitron_source.py --latest        # сүүлийнхийг татаж задлаад reports/-д
python3 bank_pipeline/capitron_source.py --dump-json capitron.json   # API хариуг хадгалах
python3 bank_pipeline/test_capitron_offline.py <Q3_workbook.xlsx>    # сүлжээгүй тест
```

Анх удаа `--list`-ээр ажиллуулж улирал зөв танигдаж байгааг шалгана уу. Танигдахгүй гарчиг
байвал `capitron.json`-оо илгээгээрэй.

## Транс банк (transbank.mn)

`bank_pipeline/transbank_source.py` — Next.js сайт. `_next/data/<buildId>/…` URL-ийн buildId сайт
шинэчлэгдэх бүрт солигддог тул кодонд хатуу бичээгүй: HTML хуудасны `__NEXT_DATA__`-аас pageProps-ийг
(эсвэл шинэ buildId-г) авна. pageProps-ийн бүтцээс үл хамааран бүх .pdf/.xlsx файлыг олж, хамгийн
ойрын объектын гарчиг / он, улирал талбараас улирлыг тодорхойлно.

```bash
python3 bank_pipeline/transbank_source.py --list
python3 bank_pipeline/transbank_source.py --dump-json transbank.json   # бүтцийг хадгалах
python3 bank_pipeline/test_transbank_offline.py <Q3_workbook.xlsx>
```

## Богд банк (bogdbank.com)

`bank_pipeline/bogdbank_source.py` — сайт жагсаалт өгдөггүй, он/улирлаар POST хийхэд
(`/report/getReportContent`, `report_type=1` = Монголбанкны стандарт) HTML хүснэгт буцаадаг. Тиймээс
workbook-ийн сүүлийн улирлаас хойш дууссан улирал бүрийг шинээс нь эхлэн асууж, хүснэгттэй хариу
ирсэн хамгийн сүүлийнхийг авна. Дүн мянган ₮ → сая ₮ автоматаар.

```bash
python3 bank_pipeline/bogdbank_source.py --probe 2026 3 --save bogd_2026q3.html
python3 bank_pipeline/test_bogdbank_offline.py <Q3_workbook.xlsx>
```

## Хувьцаа эзэмшлийн мэдээлэл (ҮЦТХТ)

Шинжилгээ хийх бүрт `bank_pipeline/ownership_source.py` нь
[Huvitsaanii-Tuvlurul](https://github.com/consultingapartment/Huvitsaanii-Tuvlurul) repo-гийн
`data/uctht_concentration.xlsx`-ийг (ҮЦТХТ-ийн сар бүрийн төвлөрлийн судалгаа) татаж
`data/ownership/`-д кэшлэнэ. Татаж чадахгүй бол кэшийг, кэш ч байхгүй бол эзэмшлийн хэсэггүйгээр үргэлжилнэ.
Банкийг ҮЦ кодоор (= МХБ-ийн компанийн дугаар 562/563/564/567/568) олж, баганыг гарчгаар нь уншина.

Шинжилгээнд нэмэгдсэн зүйлс:
- Гаргасан хувьцааны тоо ҮЦТХТ-ээс (fundamentals.json-оос зөрвөл EPS/BVPS-ийг шинэ тоогоор тооцож анхааруулна)
- Чөлөөт эргэлт (5%-иас бага эзэмшигчид) %, чөлөөт эргэлтийн зах зээлийн үнэлгээ
- Хувьцаа эзэмшигчдийг ҮЦТХТ-ийн ангиллаар ХОЁР ТУСДАА бүлгээр (хэзээ ч нийлүүлэхгүй):
  - **5%-иас ДЭЭШ** — эзэмшигчийн тоо ба түүний 3/12 сарын өөрчлөлт (ширхэгээр), эзлэх хувь ба өөрчлөлт (н.х.),
    хувьцаа, зах зээлийн үнэлгээ
  - **5%-иас ДООШ** (олон нийт, чөлөөт эргэлт) — эзэмшигчийн тоо ба 3/12 сарын өсөлт (%), эзлэх хувь,
    хувьцаа, нэг эзэмшигчид ногдох дундаж хувьцаа, зах зээлийн үнэлгээ
  - `fundamentals.json`: `holders_over5`, `pct_over5`, `holders_under5`, `pct_under5`
- Эзэмшлийн бүтцийн өөрчлөлтийн үйл явдал: давамгай хувь ≥0.5 нэгж хувь, 5%+ эзэмшигчийн тоо,
  төрийн хувь, гаргасан хувьцааны тоо, эзэмшигчдийн тоо сард ≥10%
- Word тайланд "Эзэмшлийн бүтэц ба хувьцааны хөрвөх чадвар" бүлэг, P/B–ROE–чөлөөт эргэлтийн график, Telegram мөр

```bash
python3 bank_pipeline/ownership_source.py                     # 5 банкны хураангуй + үйл явдлууд
python3 bank_pipeline/quarterly_pipeline.py ... --ownership-file local.xlsx   # локал файлаар
python3 bank_pipeline/quarterly_pipeline.py ... --no-ownership                # унтраах
python3 bank_pipeline/test_ownership_offline.py <uctht.xlsx> <Q3.xlsx> data/price_data.json data/fundamentals.json
```

## Нэмэлт харьцаа үзүүлэлт ба графикууд

`bank_pipeline/advanced_ratios.py` — шинжилгээ бүрт тооцож, Word тайлангийн "Нэмэлт харьцаа үзүүлэлт" бүлэг,
`metrics.json` ("advanced", "justified_pb", "sector_hhi") болон 11 график (`analysis/<улирал>/adv_*.png`) гаргана:

| График | Харьцаа |
|---|---|
| `adv_credit_risk.png` | NPL ба хугацаа хэтэрсэн + NPL; Texas = NPL / (ӨХ + эрсдэлийн сан) |
| `adv_stress_test.png` | ӨХ/хөрөнгө: одоо, NPL ×2, ×3 (шинэ NPL-ийн LGD хувь алдагдал) |
| `adv_funding_mix.png` | Эх үүсвэрийн бүтэц: харилцах+хадгаламж, банк/СБ, бусад эх үүсвэр, бусад |
| `adv_liquidity.png` | Хөрвөх хөрөнгө / хөрөнгө ба зээл / хадгаламж |
| `adv_dupont.png` | ROA-гийн задаргаа (дундаж хөрөнгөд): ЦХО, хүүгийн бус орлого, зардал, сан, татвар |
| `adv_earnings_buffer.png` | Сангийн өмнөх ашиг / зээл ба зээлийн эрсдэлийн өртөг |
| `adv_capital_growth.png` | Хөрөнгийн өсөлт ба дотоод капитал үүсгэлт (ROE) |
| `adv_justified_pb.png` | Бодит ба үндэслэлтэй P/B = (ROE − g)/(COE − g) |
| `adv_sector_hhi.png` | Салбарын төвлөрөл (HHI) — хөрөнгө, зээл, хадгаламж |
| `adv_trend_npl.png`, `adv_trend_roe.png` | 10 улирлын чиг хандлага, банк бүр ба салбарын медиан |

Таамаглалууд (анхдагч COE 20%, g 6%, LGD 50%) өөрчлөх: `--coe 0.18 --growth 0.05 --lgd 0.45`.
Хуудсанд мөр байхгүй бол (жиш. Голомт, Хаан банкны "хугацаа хэтэрсэн зээл") тэг биш "—" гэж харуулна.
Тест: `python3 bank_pipeline/test_advanced_ratios.py docs/Есөн_банк_шинжилгээ.xlsx`

## Word тайлангийн бүлгүүд (19 бүлэг, ~23 хуудас, 29 график)

1. Хураангуй (дүрэмд суурилсан ~20 дүгнэлт) · 2. Хэмжээ ба ашигт ажиллагаа · 3. Активын чанар, үр ашиг,
хөрвөх чадвар · 4. Зах зээлийн үнэлгээ · 5. Эзэмшлийн бүтэц (ҮЦТХТ) · 6. Банк тус бүр ·
7. Нэмэлт харьцаа: Texas, хугацаа хэтэрсэн+NPL, стресс тест, эх үүсвэр, DuPont, PPOP, ICG, үндэслэлтэй P/B, HHI ·
8. Өсөлт ба зах зээлийн хувь (YoY/QoQ, 2 жилийн CAGR, н.х. өөрчлөлт) · 9. Балансын ба орлогын бүтэц ·
10. Активын өгөөж, эх үүсвэрийн өртөг, спред, NIM · 11. Ашгийн чанар (сан/ЦХО, сан/PPOP, татварын бодит хувь,
зардал/хөрөнгө, TTM ROE, улирлын ROE, ашгийн CV) · 12. Өргөтгөсөн эрсдэл (нөөцийн харьцаа, NPL/ӨХ, цэвэр NPL/ӨХ,
зээл/эх үүсвэр, бөөний эх үүсвэр) · 13. Нэгдсэн скоркард (11 үзүүлэлтийн эрэмбэ → оноо) · 14. Хувьцааны гүйцэтгэл,
хэлбэлзэл, их уналт, өдрийн арилжаа, арилжаатай өдөр, эргэлт · 15. Өргөтгөсөн үнэлгээ (E/P, ROE/PB, Graham, PEG) ·
16. Дараагийн улирлын ашгийн чиг хандлагын тооцоо · 17. 10 улирлын чиг хандлага (NIM, зардал/орлого, ӨХ/хөрөнгө,
зээл/хадгаламж) · 18. Өгөгдлийн төлөв · 19. Графикууд

Код: `bank_analysis.py` (үндсэн), `advanced_ratios.py` (7-р бүлэг), `extended_analysis.py` (8–17-р бүлэг).
Тест: `test_advanced_ratios.py`, `test_extended_analysis.py`.

## Автомат биш хэвээр

- Ариг банк — гараар оруулна (шинжилгээнд сүүлийн огноогоор нь орно).
- Капитрон sheet-ийн "Зээлээр авсан санхүүжилт", "Дахин үнэлгээ ба бусад нөөц" мөрүүд стандарт маягтад
  шууд харгалзах мөргүй тул хоосон үлдэж болно.
- "Зохистой харьцаа" sheet (таны хүсэлтээр хөндөхгүй), CAMELS-ийн тайлбар мөрүүд, "Өмчлөл ба гүйцэтгэл"
  (master-т урьдаас буруу заасан томьёотой), Прогноз sheet.
- Word тайлангийн дүгнэлт өгүүлбэрүүд дүрэмд суурилсан — тоо нь эх тайлангаас, тайлбарыг хянаж үзнэ үү.
