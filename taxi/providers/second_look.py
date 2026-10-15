#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""النظرة الثانية على الهويّات — تُستدعى حين يشتكي الفحص من أرقام التعريف.

مأخوذة من v6.1: تُرسَل الصفحات المائلة بوضعَيها، ويُطلب من النموذج
الهويّات وحدها، لأن تكرار الصفحات في المسار العادي يُربك النموذج
في نسبة كل رقم لصاحبه.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

from taxi.core.parse import parse_json_response
from taxi.core.schema import SCHEMA
from taxi.core.enrich import NIN_FIELD, section_active, nin_fit
from taxi.pdf import pdf_to_jpegs


IDENTITY_PROMPT = """هذه صفحات من ملف سائق سيارة أجرة، أُرسلت بأوضاع متعدّدة
لأن أصلها مائل. كل صورتين متتاليتين قد تكونان الصفحة نفسها بوضعين متقابلين —
اقرأ من المعتدل وتجاهل المقلوب، ولا تحسبهما وثيقتين.

مهمّتك واحدة فقط: استخراج هويّات الأشخاص الواردين.

لكل شخص تجد له بطاقة تعريف وطنية أو شهادة ميلاد، أعد سطراً.
رقم التعريف الوطني 18 رقماً بالضبط — عُدّها مرّتين.
بنيته: خانة سلسلة · خانة جنس (0 ذكر 1 أنثى) · 3 خانات لسنة الميلاد ·
4 مكتب · 5 رقم شهادة الميلاد · 4 تسلسل.
فمولود 1979 ذكراً يبدأ رقمه بـ «10979».

أعد JSON فقط:
{"اشخاص":[{"nom_ar":"اللقب","prenom_ar":"الاسم","nin":"18 رقماً",
           "date_naissance":"YYYY-MM-DD","n_acte_naissance":"رقم الشهادة"}]}"""


def second_look_identity(pdf_path: str, provider, log=None) -> list[dict]:
    """
    قراءة ثانية موجّهة: تُرسَل الصفحات المائلة بوضعَيها، ويُطلب من النموذج
    الهويّات وحدها. تُستدعى فقط حين يشتكي الفحص من أرقام التعريف،
    لأن تكرار الصفحات في المسار العادي يُربك النموذج في نسبة كل رقم لصاحبه.
    """
    imgs = pdf_to_jpegs(pdf_path, upright=True)
    if log:
        log(f"🔍 نظرة ثانية على الهويّات — {len(imgs)} صورة بأوضاع متعدّدة")
    raw = provider.extract(imgs, IDENTITY_PROMPT)
    try:
        data = raw if isinstance(raw, dict) else parse_json_response(raw)
    except Exception:
        return []
    out = data.get("اشخاص") or data.get("persons") or []
    return [p for p in out if isinstance(p, dict)]


def merge_identity(data: dict, people: list[dict], log=None) -> list[str]:
    """
    يدمج هويّات النظرة الثانية في الأقسام التي يشكو رقمها من عطب،
    بمطابقة اللقب والاسم. لا يستبدل إلا ما كان المرشّح الجديد أفضل مطابقةً.
    """
    done: list[str] = []

    def norm(x):
        return "".join(str(x or "").split()).lower()

    for grp, ninf in NIN_FIELD.items():
        sec = data.get(grp)
        if not isinstance(sec, dict) or not section_active(data, grp):
            continue
        cur = "".join(c for c in str(sec.get(ninf) or "") if c.isdigit())
        if len(cur) == 18 and nin_fit(cur, sec) >= 2:
            continue                      # الرقم الحالي متّسق — لا نعبث به
        for p in people:
            if norm(p.get("nom_ar")) != norm(sec.get("nom_ar")):
                continue
            if sec.get("prenom_ar") and p.get("prenom_ar") and \
               norm(p["prenom_ar"]) != norm(sec["prenom_ar"]):
                continue
            new = "".join(c for c in str(p.get("nin") or "") if c.isdigit())
            if len(new) != 18 or new == cur:
                continue
            if nin_fit(new, sec) > nin_fit(cur, sec):
                sec[ninf] = new
                done.append(f"[{SCHEMA[grp]['title']}] رقم التعريف: «{cur or 'فارغ'}» "
                            f"← «{new}» (نظرة ثانية بوضع معتدل)")
                if log:
                    log(f"🔍 {done[-1]}")
            break
    return done
