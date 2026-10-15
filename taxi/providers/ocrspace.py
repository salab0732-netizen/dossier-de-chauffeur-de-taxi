#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""مزوّد OCR.space — بديل سحابي مجاني (500 طلب يومياً، 1 م.ب للصفحة).

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

from taxi.constants import WILAYA_ID, OCRSPACE_MODELS
from taxi.providers.base import ProviderError
from taxi.providers.regex_extract import extract_taxi


def _ocr_space_module():
    """تحميل ocr_space.py من مجلد البرنامج."""
    try:
        import ocr_space
        return ocr_space
    except ImportError as e:
        raise ProviderError("الملف ocr_space.py غير موجود بجانب البرنامج، أو مكتبة "
                            f"requests غير مثبّتة (pip install requests).\n{e}")


def _shrink_jpeg(data: bytes, limit: int) -> bytes:
    """تصغير صورة الصفحة حتى لا تتجاوز حدّ الحجم (1 ميغابايت في الخطة المجانية)."""
    if len(data) <= limit:
        return data
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    src = fitz.open(stream=data, filetype="jpeg")
    try:
        page = src[0]
        for zoom, quality in ((1.0, 75), (0.85, 70), (0.7, 65), (0.55, 60)):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            try:
                out = pix.tobytes("jpeg", jpg_quality=quality)
            except TypeError:                 # نسخ PyMuPDF القديمة
                out = pix.tobytes("jpeg")
            if len(out) <= limit:
                return out
    finally:
        src.close()
    return out


def _overlay_to_tsv(result: dict) -> str:
    """مواقع الكلمات من OCR.space (TextOverlay) ← صيغة TSV نفسها التي يُخرجها
    Tesseract، لتعمل قراءة خانات البطاقة الرمادية بالمواقع."""
    rows = ["level\tpage_num\tblock_num\tpar_num\tline_num\tword_num"
            "\tleft\ttop\twidth\theight\tconf\ttext"]
    for pr in result.get("ParsedResults") or []:
        lines = ((pr.get("TextOverlay") or {}).get("Lines")) or []
        for li, line in enumerate(lines, 1):
            for wi, w in enumerate(line.get("Words") or [], 1):
                txt = str(w.get("WordText", "")).replace("\t", " ").strip()
                if not txt:
                    continue
                try:
                    box = [int(float(w.get(k, 0))) for k in ("Left", "Top", "Width", "Height")]
                except (TypeError, ValueError):
                    continue
                rows.append("\t".join(["5", "1", "1", "1", str(li), str(wi),
                                       *map(str, box), "90", txt]))
    return "\n".join(rows)


class OcrSpaceProvider:
    """OCR سحابي عبر ocr.space — يدعم العربية والفرنسية/الإنجليزية معاً
    (language=auto)، ثم استخراج الحقول من النصّ بالأنماط."""
    name = "OCR.space"
    on_debug = None

    def __init__(self, key: str = "", model: str = "", *_, **__):
        self.key = (key or "").strip()
        self.model = model or OCRSPACE_MODELS[0]
        self.last_pages: list[str] = []

    def _engine(self) -> int:
        return 3 if str(self.model).endswith("3") else 2

    def _key(self, mod) -> str:
        k = self.key or str(getattr(mod, "OCR_SPACE_API_KEY", "")).strip()
        if not k or k == "ضع_مفتاحك_هنا":
            raise ProviderError("مفتاح OCR.space غير مُدخل — ضعه في تبويب الإعدادات، "
                                "أو في OCR_SPACE_API_KEY أعلى ocr_space.py.\n"
                                "مفتاح مجاني: ocr.space/ocrapi/freekey")
        return k

    def check(self) -> str:
        mod = _ocr_space_module()
        key = self._key(mod)
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        # صورة صغيرة فيها كلمة واحدة — يستهلك طلباً واحداً من الحصة
        doc = fitz.open()
        pg = doc.new_page(width=200, height=60)
        pg.insert_text((20, 40), "TEST 32", fontsize=20)
        img = pg.get_pixmap().tobytes("png")
        doc.close()
        try:
            text = mod.ocr_space_bytes(img, "check.png", api_key=key, engine=self._engine())
        except mod.OcrSpaceError as e:
            raise ProviderError(str(e))
        seen = f" — قرأ «{text.strip()[:20]}»" if text.strip() else ""
        return f"OCR.space ✓ — المحرّك {self._engine()}{seen}"

    def list_models(self) -> list[str]:
        return list(OCRSPACE_MODELS)

    def extract(self, images: list[bytes], prompt: str) -> str:
        mod = _ocr_space_module()
        key = self._key(mod)
        limit = int(getattr(mod, "MAX_FILE_BYTES", 1024 * 1024))
        texts: list[str] = []
        tsvs: list[str] = []
        for i, img in enumerate(images, 1):
            data = _shrink_jpeg(img, limit)
            try:
                res = mod.ocr_space_request(data, f"page{i}.jpg", api_key=key,
                                            engine=self._engine(), overlay=True)
            except (mod.OcrSpaceAuthError, mod.OcrSpaceQuotaError,
                    mod.OcrSpaceConnectionError) as e:
                # المفتاح والحصّة والاتصال تخصّ كل الصفحات — نوقف الملف
                raise ProviderError(f"الصفحة {i}/{len(images)}: {e}")
            except mod.OcrSpaceError as e:
                # صفحة واحدة غير مقروءة لا تُسقط الملف كله
                if self.on_debug:
                    self.on_debug(f"⚠ الصفحة {i}: {e}")
                texts.append("")
                tsvs.append("")
                continue
            page_text = "\n".join((pr.get("ParsedText") or "").replace("\r\n", "\n")
                                  for pr in res.get("ParsedResults") or []
                                  if pr.get("FileParseExitCode") == 1)
            texts.append(page_text.replace("\f", "\n"))
            tsvs.append(_overlay_to_tsv(res))
            if self.on_debug:
                self.on_debug(f"OCR.space: الصفحة {i}/{len(images)} ✓ "
                              f"({len(data) // 1024} ك.ب)")
        if not any(t.strip() for t in texts):
            raise ProviderError("لم يُستخرج أي نص من الملف عبر OCR.space.")

        # النصّ الخام يُحفَظ ليُعرَض في شاشة المراجعة إلى جانب الصورة.
        # بدونه لا سبيل لمعرفة سبب خانةٍ فارغة: أهو أن المسح لم يقرأ
        # القيمة، أم أنها قُرئت والمستخرج لم يلتقطها؟ والفرق بينهما
        # هو الفرق بين إصلاحٍ في محلّه وتخمينٍ في الهواء.
        self.last_pages = texts
        return json.dumps(extract_taxi("\f".join(texts), WILAYA_ID, tsv="\f".join(tsvs)),
                          ensure_ascii=False)
