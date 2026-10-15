#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""تطبيع القيم: الأرقام العربية، تصحيح حروف المسح، التواريخ، القيم المحصورة.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import datetime
import re

from taxi.core.schema import SCHEMA, ENUM_FIELDS
from taxi.core.text import _norm
from taxi.core.vocabulary import _VOCAB
from taxi.core.enrich import repair_nin


DIGIT_FIELDS = {"nin", "nin_modjahid"}
# خلط شائع في المسح الضوئي: حرف ↔ رقم
OCR_DIGITS = str.maketrans({"O": "0", "o": "0", "Q": "0", "D": "0",
                            "I": "1", "l": "1", "|": "1",
                            "S": "5", "B": "8", "Z": "2", "G": "6"})


def normalize(data: dict, log=None) -> list[str]:
    """
    يُصلح أخطاء القراءة الضوئية الواضحة قبل الفحص.
    يعدّل data في مكانه ويعيد قائمة بما غُيّر.
    لا يصلح إلا ما كان تصحيحه قاطعاً — ما بقي ملتبساً يذهب للمراجعة.
    """
    fixed: list[str] = []

    def note(grp, fkey, old, new, why):
        label = next((lb for f, lb, _ in SCHEMA[grp]["fields"] if f == fkey), fkey)
        fixed.append(f"[{SCHEMA[grp]['title']}] {label}: «{old}» ← «{new}» ({why})")
        if log:
            log(f"🔧 {fixed[-1]}")

    for gkey, grp in SCHEMA.items():
        sec = data.get(gkey)
        if not isinstance(sec, dict):
            continue
        for fkey, _, _ in grp["fields"]:
            v = sec.get(fkey)
            if v in (None, "", "null"):
                continue
            s = str(v).strip()
            orig = s

            # ① حقول رقمية بحتة: حروف شائعة الخلط ← أرقام
            if fkey in DIGIT_FIELDS:
                t = s.replace(" ", "").replace("-", "").replace(".", "")
                t = t.translate(OCR_DIGITS)
                if t.isdigit() and t != orig:
                    sec[fkey] = t
                    note(gkey, fkey, orig, t, "تصحيح قراءة ضوئية")
                    s = t
                # محاولة إصلاح الطول باستعمال بقية الحقول كمرجع
                if t.isdigit() and len(t) == 19:
                    dn = sec.get("date_naissance")
                    yr = None
                    if dn:
                        try:
                            yr = datetime.date.fromisoformat(str(dn)[:10]).year
                        except ValueError:
                            pass
                    ac = sec.get("n_acte_naissance")
                    if ac is not None and not str(ac).strip().isdigit():
                        ac = None
                    rep, why = repair_nin(t, sec.get("sexe"), yr, ac)
                    if rep:
                        sec[fkey] = rep
                        note(gkey, fkey, t, rep, why)
                        s = rep

            # ②a تاريخ من سنة فقط — معلومة ناقصة لا تاريخ
            if "date" in fkey and s.isdigit() and len(s) == 4:
                sec[fkey] = None
                note(gkey, fkey, orig, "فارغ", "سنة بلا يوم ولا شهر")
                continue

            # ② تواريخ مشوّهة: يوم أو شهر = 00 يعني أنه غير مقروء
            if "date" in fkey and len(s) >= 7:
                if s[4:5] == "-" and ("-00" in s[4:] or s.endswith("-0")):
                    sec[fkey] = None
                    note(gkey, fkey, orig, "فارغ", "يوم أو شهر غير مقروء")
                    continue
                try:
                    datetime.date.fromisoformat(s[:10])
                except ValueError:
                    pass

            # ③ قيم محصورة: مطابقة متسامحة بالمسافات والحالة
            allowed = (_VOCAB.get(gkey) or {}).get(fkey)
            if allowed and s not in allowed:
                def squash(x):
                    return "".join(c for c in str(x).lower() if c.isalnum())
                hit = next((a for a in allowed if squash(a) == squash(s)), None)
                if hit:
                    sec[fkey] = hit
                    note(gkey, fkey, orig, hit, "توحيد مع المرجعية")

    # ④ رقم الباب: انسخه إلى الأقسام التي أغفلته
    portes = {str(data.get(g, {}).get(SCHEMA[g]["porte"])).strip()
              for g in SCHEMA if SCHEMA[g].get("porte")
              and (data.get(g) or {}).get(SCHEMA[g]["porte"]) not in (None, "", "null")}
    if len(portes) == 1:
        p = portes.pop()
        for g in SCHEMA:
            pf = SCHEMA[g].get("porte")
            if not pf:
                continue
            sec = data.setdefault(g, {})
            if sec.get(pf) in (None, "", "null"):
                sec[pf] = p
                note(g, pf, "فارغ", p, "منسوخ من قسم آخر")

    return fixed
