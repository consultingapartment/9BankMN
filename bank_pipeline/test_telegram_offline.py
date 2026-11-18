#!/usr/bin/env python3
"""
test_telegram_offline.py — Telegram-д ЗӨВХӨН Word файл илгээгдэхийг шалгана (requests-ийг дуурайлгана).
    python3 test_telegram_offline.py <analysis/2026Q3 хавтас>
"""
import os
import sys
import types
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import quarterly_pipeline as qp  # noqa: E402

FAIL = []


def check(c, m):
    print(("  ✓ " if c else "  ✗ ") + m)
    if not c:
        FAIL.append(m)


class Resp:
    def __init__(self, code=200):
        self.status_code, self.text = code, "x"


def fake(calls, codes):
    m = types.ModuleType("requests")

    def post(url, data=None, files=None, timeout=30):
        calls.append({"method": url.rsplit("/", 1)[-1], "data": data or {}, "files": files or {}})
        return Resp(codes.pop(0) if codes else 200)
    m.post = post
    return m


def main(an):
    doc = next(Path(an).glob("*.docx"))
    os.environ["TELEGRAM_BOT_TOKEN"], os.environ["TELEGRAM_CHAT_ID"] = "T", "C"
    calls = []
    sys.modules["requests"] = fake(calls, [])
    ok = qp.tg_send_document(doc)
    check(ok and [c["method"] for c in calls] == ["sendDocument"], f"зөвхөн sendDocument: {[c['method'] for c in calls]}")
    check(set(calls[0]["data"]) == {"chat_id"}, "тайлбар (caption), текст байхгүй")
    check(calls[0]["files"]["document"][0] == doc.name, f"файл: {doc.name}")
    calls = []
    sys.modules["requests"] = fake(calls, [500, 200])
    check(qp.tg_send_document(doc) and len(calls) == 2, "алдаа гарвал нэг удаа дахин оролдоно")
    calls = []
    sys.modules["requests"] = fake(calls, [])
    check(not qp.tg_send_document(Path("/nope.docx")) and not calls, "файл байхгүй бол юу ч илгээхгүй")
    del os.environ["TELEGRAM_BOT_TOKEN"]
    check(not qp.tg_send_document(doc) and not calls, "токенгүй бол илгээхгүй")
    src = (HERE / "quarterly_pipeline.py").read_text(encoding="utf-8")
    check("sendMessage" not in src and "sendMediaGroup" not in src, "pipeline-д текст/цомог илгээх код үлдээгүй")
    print("\n" + ("✅ БҮХ ТЕСТ АМЖИЛТТАЙ" if not FAIL else f"❌ {len(FAIL)}: " + "; ".join(FAIL)))
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
