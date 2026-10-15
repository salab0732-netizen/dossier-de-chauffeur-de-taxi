#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""خطّ المعالجة: من ملف PDF إلى سجلّ مُتحقَّق منه.

يجمع المراحل التي كانت موزّعة في دالة الواجهة القديمة، فيصبح الخطّ
قابلاً للاستدعاء من الخادم ومن سطر الأوامر ومن الاختبارات سواءً.

المراحل:
  1. PDF ← صور          (الصفحات الأفقية بوضعَين، فهي بطاقات مائلة)
  2. صور ← JSON         (المزوّد: نموذج سحابي أو OCR محلي)
  3. تطبيع              (أرقام عربية، حروف المسح، تواريخ، قيم محصورة)
  4. استنتاج            (الثوابت، النشر بين الأقسام، إصلاح NIN)
  5. فحص                (منطقي — لا يُدخَل إلا ما اجتازه بصفر تنبيهات)
  6. نظرة ثانية         (إن شكا الفحص من أرقام التعريف وحدها)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from taxi.core.enrich import enrich
from taxi.core.normalize import normalize
from taxi.core.parse import parse_json_response
from taxi.core.prompt import build_prompt
from taxi.core.validate import validate
from taxi.core.vocabulary import learn_vocabulary
from taxi.pdf import pdf_to_jpegs
from taxi.providers.base import Cancelled
from taxi.providers.second_look import merge_identity, second_look_identity


@dataclass
class Result:
    """حصيلة معالجة ملف واحد."""
    data: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    pages: int = 0

    @property
    def clean(self) -> bool:
        """صفر تنبيهات — يُدخَل تلقائياً."""
        return not self.warnings


def run(pdf_path: str, provider, excel_path: str = "", *,
        log=None, progress=None, should_stop=None,
        second_look: bool = True) -> Result:
    """يُعالج ملفاً واحداً ويعيد السجلّ مع التنبيهات.

    log      : رسالة نصّية للموظّف
    progress : (المرحلة، 0..1) لشريط التقدّم
    """
    def say(m, kind="info"):
        if log:
            log(m, kind)

    def step(stage, frac=0.0):
        if progress:
            progress(stage, frac)

    def check():
        if should_stop and should_stop():
            raise Cancelled()

    res = Result()

    # ── المصطلحات من الملف نفسه، قبل بناء التوجيه ──
    if excel_path and Path(excel_path).exists():
        step("تعلّم المصطلحات", 0.02)
        vocab = learn_vocabulary(excel_path)
        if vocab:
            say(f"تُعلِّمت {len(vocab)} حقلاً محصوراً من ورقة Referances.", "dim")

    # ── 1. صور ──
    check()
    step("تحويل الصفحات", 0.05)
    images = pdf_to_jpegs(pdf_path, upright=True,
                          progress=lambda i, n: step("تحويل الصفحات",
                                                     0.05 + 0.15 * i / max(n, 1)))
    res.pages = len(images)
    say(f"{len(images)} صورة من الملف.", "dim")

    # ── 2. استخراج ──
    check()
    step("الاستخراج", 0.25)
    raw = provider.extract(images, build_prompt())
    res.data = raw if isinstance(raw, dict) else parse_json_response(raw)
    say("استُخرجت البيانات.", "ok")

    # ── 3-4. تطبيع واستنتاج ──
    check()
    step("التطبيع", 0.70)
    res.notes += normalize(res.data, log=lambda m: say(m, "dim"))
    step("الاستنتاج", 0.78)
    res.notes += enrich(res.data, log=lambda m: say(m, "dim"))

    # ── 5. فحص ──
    step("الفحص", 0.85)
    res.warnings = validate(res.data)

    # ── 6. نظرة ثانية — فقط إن كانت الشكوى في أرقام التعريف ──
    nin_trouble = any(("رقم التعريف" in w) or ("NIN" in w) for w in res.warnings)
    if second_look and nin_trouble and hasattr(provider, "extract"):
        check()
        step("نظرة ثانية", 0.90)
        say("أرقام التعريف مشكوك فيها — نظرة ثانية بوضع معتدل.", "warn")
        try:
            people = second_look_identity(pdf_path, provider,
                                          log=lambda m: say(m, "dim"))
            merged = merge_identity(res.data, people, log=lambda m: say(m, "ok"))
            if merged:
                say(f"صُحّح {len(merged)} رقم تعريف.", "ok")
                res.notes += merged
                normalize(res.data)
                enrich(res.data)
                res.warnings = validate(res.data)
            else:
                say("النظرة الثانية لم تجد أفضل ممّا لدينا.", "dim")
        except Cancelled:
            raise
        except Exception as e:
            say(f"تعذّرت النظرة الثانية: {e}", "warn")

    step("تمّ", 1.0)
    if res.clean:
        say("اجتاز الفحص بصفر تنبيهات.", "ok")
    else:
        say(f"{len(res.warnings)} تنبيهاً — بانتظار المراجعة.", "warn")
    return res
