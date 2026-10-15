#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""مزوّد Offline — Tesseract OCR محلي بالكامل، بدون إنترنت ولا مفتاح API.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from taxi.constants import WILAYA_ID, OFFLINE_MODELS
from taxi.providers.base import ProviderError
from taxi.providers.regex_extract import extract_taxi


OFFLINE_MODELS = ["tesseract-ara", "tesseract-ara+fra"]
TESSDATA_DIRS = [
    # مسار ASCII بجانب Documents (بدون أحرف عربية في المسار)
    str(Path(__file__).resolve().parent.parent / "taxi_tessdata"),
    # مجلد tessdata بجانب البرنامج (قد يفشل إن كان المسار عربياً)
    str(Path(__file__).resolve().parent / "tessdata"),
    r"C:\Program Files\Tesseract-OCR\tessdata",
    r"C:\Program Files (x86)\Tesseract-OCR\tessdata",
    str(Path(sys.executable).parent / "tessdata"),
]


def _find_tesseract() -> str:
    """يبحث عن مسار tesseract.exe."""
    import shutil
    found = shutil.which("tesseract")
    if found:
        return found
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return "tesseract"          # نأمل أن يكون في PATH


def _find_tessdata() -> str | None:
    """يعيد أول مجلد tessdata يحتوي ara.traineddata."""
    for d in TESSDATA_DIRS:
        if Path(d, "ara.traineddata").exists():
            return d
    return None


class OfflineProvider:
    """استخراج محلي 100 % — Tesseract OCR + Regex، بدون إنترنت ولا مفتاح API."""
    name  = "Offline"
    model = "tesseract-ara"
    on_debug = None

    def __init__(self, *_, **__):
        self._tess     = _find_tesseract()
        self._tessdata = _find_tessdata()

    # ── فحص الجاهزية ────────────────────────────────────────────────
    def check(self) -> str:
        import subprocess
        try:
            r = subprocess.run([self._tess, "--version"],
                               capture_output=True, text=True, timeout=10)
        except FileNotFoundError:
            raise ProviderError(
                "Tesseract غير مثبّت.\n"
                "حمّله من: https://github.com/UB-Mannheim/tesseract/wiki")
        if not self._tessdata:
            raise ProviderError(
                "ملف ara.traineddata غير موجود.\n"
                "انقر زر ⬇ تثبيت اللغة في تبويب الإعدادات، أو ضع\n"
                "ara.traineddata داخل مجلد tessdata/ بجانب البرنامج.")
        ver = r.stdout.splitlines()[0] if r.stdout else "Tesseract ✓"
        return f"{ver} | tessdata: {self._tessdata}"

    def list_models(self) -> list[str]:
        return list(OFFLINE_MODELS)

    # ── OCR الفعلي ──────────────────────────────────────────────────
    def extract(self, images: list[bytes], prompt: str) -> str:
        import subprocess, tempfile, os

        import shutil as _shutil

        # تحديث tessdata عند كل استدعاء (قد يكون ثُبِّت للتو)
        self._tessdata = _find_tessdata()
        if not self._tessdata:
            raise ProviderError(
                "ملف ara.traineddata غير موجود.\n"
                "الحل: انقر زر ⬇ تثبيت اللغة العربية في تبويب الإعدادات،\n"
                "أو شغّل:  python offline_ocr.py --setup\n"
                "أو انسخ ara.traineddata إلى مجلد tessdata/ بجانب البرنامج.")

        # Tesseract على Windows لا يدعم مسارات Unicode — ننسخ لـ temp إذا لزم
        if any(ord(c) > 127 for c in self._tessdata):
            _safe = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")),
                                 "taxi_tessdata")
            os.makedirs(_safe, exist_ok=True)
            for _lng in ("ara", "fra"):
                _src = os.path.join(self._tessdata, f"{_lng}.traineddata")
                _dst = os.path.join(_safe, f"{_lng}.traineddata")
                if os.path.exists(_src) and not os.path.exists(_dst):
                    _shutil.copy2(_src, _dst)
            self._tessdata = _safe

        # الفرنسية تضيف الأسماء اللاتينية والصنف والطراز والرقم التسلسلي
        # (البطاقة الرمادية وشهادات CNAS/CASNOS) — تُستعمل تلقائياً إن وُجد ملفها
        has_fra = os.path.exists(os.path.join(self._tessdata, "fra.traineddata"))
        lang = "ara+fra" if has_fra else "ara"
        if self.on_debug:
            self.on_debug(f"OCR: {lang}" + ("" if has_fra else
                          "  (أضف fra.traineddata لقراءة الحقول اللاتينية)"))

        pages_text: list[str] = []
        pages_tsv: list[str] = []

        for img_bytes in images:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as fh:
                fh.write(img_bytes)
                tmp_in = fh.name

            tmp_out = tmp_in[:-4]   # بدون .jpg
            # تمريرة واحدة تُخرج النص (.txt) ومواقع الكلمات (.tsv) معاً — بلا وقت إضافي.
            # المواقع تُستعمل لقراءة خانات جدول البطاقة الرمادية.
            cmd = [self._tess, tmp_in, tmp_out, "-l", lang, "--psm", "6",
                   "-c", "tessedit_create_txt=1", "-c", "tessedit_create_tsv=1"]
            if self._tessdata:
                cmd += ["--tessdata-dir", self._tessdata]

            try:
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                if self.on_debug and r.stderr.strip():
                    self.on_debug(r.stderr.strip()[:120])
            except subprocess.TimeoutExpired:
                # صفحة صورة/ضجيج تستغرق طويلاً — نتجاوزها بدل إسقاط الملف كله
                if self.on_debug:
                    self.on_debug(f"⏱ تجاوز الصفحة {len(pages_text) + 1} (OCR > 120 ثانية)")

            for ext, bucket in ((".txt", pages_text), (".tsv", pages_tsv)):
                fpath = tmp_out + ext
                if os.path.exists(fpath):
                    with open(fpath, "r", encoding="utf-8") as fh:
                        bucket.append(fh.read().replace("\f", "\n"))
                    os.remove(fpath)
                else:
                    bucket.append("")
            try:
                os.remove(tmp_in)
            except OSError:
                pass

        # الصفحات مفصولة بـ \f: الاستخراج يتعرّف على نوع كل وثيقة على حدة
        return self._build_json("\f".join(pages_text), "\f".join(pages_tsv))

    # ── تحليل النص إلى JSON ─────────────────────────────────────────
    @staticmethod
    def _rx(text: str, *patterns, default=None):
        import re
        for p in patterns:
            m = re.search(p, text, re.MULTILINE)
            if m:
                return m.group(1).strip()
        return default

    def _build_json(self, text: str, tsv: str | None = None) -> str:
        """النص (صفحات مفصولة بـ \\f) ← JSON بأقسام المخطط نفسها:
        operateur · chauffeur · vehicule · licence_modjahid."""
        return json.dumps(extract_taxi(text, WILAYA_ID, tsv=tsv), ensure_ascii=False)
