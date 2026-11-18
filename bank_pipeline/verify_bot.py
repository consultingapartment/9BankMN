#!/usr/bin/env python3
"""
verify_bot.py — bank_financials_bot.py ЗӨВ ажиллаж байгааг Q3 гараагүй
байхад ч БОДИТ сүлжээ, БОДИТ өгөгдлөөр шалгах скрипт.

АРГА ЗҮЙ (rollback-and-rederive):
  Q3 2026 тайлан гараагүй ч Q2 2026 (2026/06/30) аль хэдийн МАСТЕР
  ФАЙЛД байгаа бөгөөд open.mse.mn дээр ч НИЙТЛЭГДСЭН хэвээр байгаа.
  Иймд:
    1. Мастер файлын хуулбар дээр 5 банкны сүүлийн огнооны THЛГОЙГ
       ЗӨВХӨН ("2026/06/30" -> "2026/03/31") ХУУРАМЧААР ХУУЧИРУУЛНА
       (утгыг хусахгүй, зөвхөн толгойн текстийг өөрчилнө).
    2. bank_financials_bot.py-г ЭНЭ хуулбар дээр жинхэнэ сүлжээгээр
       ажиллуулна — бот "сүүлийн мэдээлэл 2026/03/31 байна" гэж
       андуурч, Q2 (2026/06/30)-г ШИНЭ гэж үзээд, ЖИНХЭНЭ
       open.mse.mn-ээс дахин татаж, шинэ багана болгон бичнэ.
    3. Ингэж шинээр бичигдсэн баганыг ХУУЛБАРЛАГДААГҮЙ эх файлын
       ЖИНХЭНЭ 2026/06/30 баганатай харьцуулна — АЯНДАА мэдэгдэж буй
       зөв хариулттай тул бот зөв/буруу ажилласныг тодорхой хардаг.

Ажиллуулах:
    python3 verify_bot.py Есөн_банк_шинжилгээ2026-06-30.xlsx
"""
import sys
import shutil
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).parent))
import bank_financials_bot as bot


def make_rolled_back_copy(original_path: str, test_path: str):
    """5 банкны баланс/орлого-ийн СҮҮЛИЙН огнооны толгойг л хуучин
    огноо болгож 'хуурна' — утгыг хусахгүй, зөвхөн шошгыг өөрчилнө."""
    shutil.copy(original_path, test_path)
    wb = load_workbook(test_path)
    for bank in bot.BANKS:
        for kind, hdr in (("баланс", bot.HEADER_ROW[bank]["bal"]),
                          ("орлого", bot.HEADER_ROW[bank]["inc"])):
            ws = wb[f"{bank} {kind}"]
            max_col = ws.max_column
            real_date = str(ws.cell(hdr, max_col).value).strip()
            print(f"  {bank} {kind}: '{real_date}' -> '2026/03/31' (толгой л, утга хэвээр)")
            ws.cell(hdr, max_col, value="2026/03/31")
    # Зохистой харьцааны толгойг ч мөн адил "хуучруулна" (5 блокийн бүгд)
    if bot.RATIO_SHEET in wb.sheetnames:
        ws = wb[bot.RATIO_SHEET]
        max_col = ws.max_column
        for hdr_row, _ in bot.RATIO_BLOCKS:
            ws.cell(hdr_row, max_col, value="2026/03")
    wb.save(test_path)


def compare_results(original_path: str, test_path: str):
    """Шинээр (сүлжээгээр) бичигдсэн Q2 баганыг эх файлын ЖИНХЭНЭ Q2
    баганатай харьцуулж, ялгааг хувиар хэвлэнэ."""
    orig_wb = load_workbook(original_path, data_only=True)
    test_wb = load_workbook(test_path, data_only=True)

    print("\n" + "=" * 70)
    print("ХАРЬЦУУЛАЛТ: бодит эх файл дэх Q2 vs бот дахин татаж бичсэн Q2")
    print("=" * 70)

    all_ok = True
    for bank in bot.BANKS:
        for kind, hdr in (("баланс", bot.HEADER_ROW[bank]["bal"]),
                          ("орлого", bot.HEADER_ROW[bank]["inc"])):
            orig_ws = orig_wb[f"{bank} {kind}"]
            test_ws = test_wb[f"{bank} {kind}"]
            orig_col = orig_ws.max_column
            # шинэ файлд Q2 нь өмнөх "2026/03/31"-ийн ДАРАА нэмэгдсэн багана
            test_col = test_ws.max_column

            mismatches = []
            for r in range(hdr + 1, orig_ws.max_row + 1):
                label = orig_ws.cell(r, 1).value
                if not label:
                    continue
                orig_val = orig_ws.cell(r, orig_col).value
                test_val = test_ws.cell(r, test_col).value
                if orig_val is None and test_val is None:
                    continue
                if orig_val is not None and test_val is None:
                    # Эх файлд ЖИНХЭНЭ утга байсан атлаа бот дахин
                    # бичихдээ юу ч тааруулж чадаагүй — энэ бол БОДИТ
                    # алдаа, "тохирол олоогүй учир хэвийн" гэж хуурамчаар
                    # алгасаж БОЛОХГҮЙ (өмнө нь яг ийм алгасалт "хуурмаг
                    # ногоон" илэрц үзүүлж байсныг олсон).
                    mismatches.append((label, orig_val, "БИЧИГДЭЭГҮЙ (None)", 100.0))
                    continue
                if orig_val is None and test_val is not None:
                    continue  # бот шинэ (эхэндээ хоосон байсан) мөрд нэмэлт утга бичсэн — асуудалгүй
                if isinstance(orig_val, (int, float)) and isinstance(test_val, (int, float)):
                    if orig_val == 0:
                        continue
                    diff_pct = abs(test_val - orig_val) / abs(orig_val) * 100
                    if diff_pct > 1.0:  # >1% зөрүү — жинхэнэ анхаарал татах зүйл
                        mismatches.append((label, orig_val, test_val, diff_pct))

            if mismatches:
                all_ok = False
                print(f"\n❌ {bank} {kind}: {len(mismatches)} мөрөнд >1% зөрүүтэй")
                for label, ov, tv, dp in mismatches:
                    ov_str = f"{ov:,.1f}" if isinstance(ov, (int, float)) else str(ov)
                    tv_str = f"{tv:,.1f}" if isinstance(tv, (int, float)) else str(tv)
                    print(f"     {label}: эх={ov_str}  бот={tv_str}  зөрүү={dp:.1f}%")
            else:
                print(f"✓ {bank} {kind}: бүх тааруулагдсан мөр таарлаа")

    print("\n" + "=" * 70)
    if all_ok:
        print("✅ АМЖИЛТТАЙ — бот жинхэнэ сүлжээгээр татаж бичсэн утга нь "
              "эх файлын мэдэгдэж буй зөв утгатай таарч байна.")
    else:
        print("⚠️ ЗӨРҮҮ ОЛДЛОО — дээрх мөрүүдийг нягталж, шаардлагатай бол "
              "bank_financials_bot.py-г засварлана уу.")
    print("=" * 70)


def main():
    # bank_financials_bot.py-д аль хэдийн зассан Jupyter/Colab-д зориулсан
    # найдвартай замын тодорхойлолтыг дахин ашиглана (sys.argv-д Jupyter-ийн
    # өөрийн "-f /root/.../kernel-xxx.json" орж ирдэг асуудлыг шийднэ).
    original_path = bot.resolve_excel_path()
    if original_path is None:
        print("Ашиглах:")
        print("  Терминалаас:  python3 verify_bot.py <Есөн_банк_шинжилгээ....xlsx>")
        print("  Colab/Jupyter-ээс:")
        print('    !python3 verify_bot.py "Есөн_банк_шинжилгээ2026-06-30.xlsx"')
        print("  Эсвэл: файлаа скрипттэй ЯГ НЭГ хавтаст байрлуулаад аргументгүйгээр ажиллуулж болно.")
        sys.exit(1)

    test_path = "TEST_" + Path(original_path).name

    print("1-Р АЛХАМ: хуулбар үүсгэж, сүүлийн огнооны ТОЛГОЙГ л хуучруулж байна...")
    make_rolled_back_copy(original_path, test_path)

    print(f"\n2-Р АЛХАМ: bank_financials_bot.py-г '{test_path}' дээр ЖИНХЭНЭ "
          f"сүлжээгээр ажиллуулж байна...")
    print("   (энэ нь open.mse.mn-ээс Q2 2026 өгөгдлийг ЖИНХЭНЭ дахин татна)\n")
    sys.argv = ["bank_financials_bot.py", test_path]
    bot.main()

    compare_results(original_path, test_path)
    print(f"\nℹ Тест файл '{test_path}'-г шалгаж дууссаны дараа устгаж болно.")


if __name__ == "__main__":
    main()
