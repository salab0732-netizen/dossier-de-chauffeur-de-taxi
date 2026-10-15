#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""تحويل صفحات PDF إلى صور.

من v6.1: الصفحات الأفقية غالباً بطاقاتٌ مُسحت ضوئياً وهي مائلة،
وقراءتها مائلةً تُفسد الأرقام الطويلة. فتُرسَل بوضعَيها المحتملَين
(90° يميناً و90° يساراً) ليقرأ النموذج من المعتدل منهما.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

from taxi.providers.base import Cancelled


def pdf_to_jpegs(pdf_path: str, zoom: float = 2.0, progress=None,
                 upright: bool = False) -> list[bytes]:
    """
    تحويل صفحات PDF إلى صور JPEG.

    الصفحات الأفقية غالباً بطاقاتٌ مُسحت ضوئياً وهي مائلة، وقراءتها
    مائلةً تُفسد الأرقام الطويلة. فتُرسَل بوضعَيها المحتملَين (90° يميناً
    و90° يساراً) ليقرأ النموذج من المعتدل منهما.
    تحقّقٌ عملي: رقم تعريف أخطأ فيه النموذج مراراً وهو مائل، أصابه
    حرفاً بحرف حين أُرسل معتدلاً. والأفقية 15% من الصفحات فقط.
    """
    try:
        import pymupdf as fitz           # الاسم الحديث
    except ImportError:
        import fitz                      # الاسم القديم
    doc = fitz.open(pdf_path)
    out: list[bytes] = []
    total = len(doc)
    try:
        for i, page in enumerate(doc):
            base = fitz.Matrix(zoom, zoom)
            landscape = page.rect.width > page.rect.height
            if upright and landscape:
                for ang in (90, -90):
                    out.append(page.get_pixmap(
                        matrix=base * fitz.Matrix(1, 1).prerotate(ang)
                    ).tobytes("jpeg"))
            else:
                out.append(page.get_pixmap(matrix=base).tobytes("jpeg"))
            if progress:
                progress(i + 1, total)
    finally:
        doc.close()
    return out


# ══════════════════════════════════════════════════════════════════════
#  مزوّدو الذكاء الاصطناعي
# ══════════════════════════════════════════════════════════════════════
