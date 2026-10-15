#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""الفحص المنطقي قبل الكتابة — لا يُدخَل تلقائياً إلا ما اجتازه بصفر تنبيهات.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import datetime
import re

from taxi.constants import WILAYA_ID
from taxi.core.schema import SCHEMA, ENUM_FIELDS, EXPECTED, SOURCE_DOC
from taxi.core.vocabulary import _VOCAB
from taxi.core.text import _norm
from taxi.core.enrich import (section_active, decode_nin, decode_cnas,
                              nin_fit, NIN_FIELD,
                              detect_swapped_nin, gather_witnesses)


def validate(data: dict) -> list[str]:
    """فحص منطقي للبيانات المستخرجة — يعيد قائمة تنبيهات."""
    warn: list[str] = []

    def val(grp, key):
        v = (data.get(grp) or {}).get(key)
        return None if v in (None, "", "null") else str(v).strip()

    # رقم الباب متطابق (حقل الباب، لا مفتاح كشف التكرار)
    portes = {g: val(g, SCHEMA[g].get("porte"))
              for g in SCHEMA if SCHEMA[g].get("porte")}
    seen = {v for v in portes.values() if v}
    if not seen:
        warn.append("رقم الباب غير موجود في أي قسم")
    elif len(seen) > 1:
        warn.append(f"رقم الباب غير متطابق بين الأقسام: {', '.join(sorted(seen))}")

    # NIN = 18 رقم
    for grp, key, who in (("operateur", "nin", "صاحب الرخصة"),
                          ("chauffeur", "nin", "السائق"),
                          ("licence_modjahid", "nin_modjahid", "المجاهد")):
        v = val(grp, key)
        if v:
            digits = "".join(ch for ch in v if ch.isdigit())
            if len(digits) != 18:
                warn.append(f"NIN {who} ليس 18 رقماً ({len(digits)})")

    # التواريخ
    for grp in SCHEMA:
        for fkey, flabel, _ in SCHEMA[grp]["fields"]:
            if not fkey.startswith("date") and "date" not in fkey:
                continue
            v = val(grp, fkey)
            if v:
                try:
                    datetime.date.fromisoformat(v[:10])
                except ValueError:
                    warn.append(f"[{SCHEMA[grp]['title']}] {flabel}: تاريخ غير صالح «{v}»")

    # صلاحية الرخصة
    d1, d2 = val("operateur", "date_debut_contrat"), val("operateur", "date_fin_contrat")
    if d1 and d2:
        try:
            if datetime.date.fromisoformat(d1[:10]) >= datetime.date.fromisoformat(d2[:10]):
                warn.append("تاريخ بداية الرخصة بعد تاريخ نهايتها")
        except ValueError:
            pass

    # حقول أساسية
    for grp, key, label in (("operateur", "nom_ar", "لقب المتعامل"),
                            ("vehicule", "n_immatriculation", "رقم تسجيل المركبة")):
        if not val(grp, key):
            warn.append(f"حقل أساسي فارغ: {label}")

    # قيم خارجة عن أعراف الملف
    for gkey, flds in (_VOCAB or {}).items():
        for fkey, allowed in flds.items():
            v = val(gkey, fkey)
            if v and v not in allowed:
                label = next((lb for f, lb, _ in SCHEMA[gkey]["fields"] if f == fkey), fkey)
                warn.append(f"[{SCHEMA[gkey]['title']}] {label}: «{v}» ليست من قيم الملف "
                            f"({' | '.join(allowed[:4])})")

    # المتعامل وذو الحق يجب أن يكونا شخصين مختلفين عادةً
    o, l = val("operateur", "nin"), val("licence_modjahid", "nin_modjahid")
    if o and l and o == l:
        warn.append("المتعامل وذو الحق لهما نفس رقم التعريف — تحقّق من عدم الخلط بينهما")

    # الرقم التسلسلي للمركبة لا يكون 18 رقماً بحتاً — تلك بصمة رقم تعريف وطني
    vin = val("vehicule", "n_chasis")
    if vin and vin.isdigit() and len(vin) == 18:
        d = decode_nin(vin)
        extra = ""
        if d.get("sexe"):
            extra = (f" (بنيته سليمة: {d['sexe']}، مواليد "
                     f"{d.get('annee_naissance','؟')}، شهادة {d.get('n_acte_naissance','؟')})")
        warn.append(f"[المركبة] الرقم التسلسلي «{vin}» من 18 رقماً بحتاً — "
                    f"هذه بنية رقم تعريف وطني لا رقم تسلسلي{extra}")

    # رقم شهادة ميلاد غير معقول (القاعدة: 5 خانات كحدّ أقصى)
    for grp in SCHEMA:
        v = val(grp, "n_acte_naissance")
        if v and v.isdigit() and len(v) > 5:
            warn.append(f"[{SCHEMA[grp]['title']}] رقم شهادة الميلاد «{v}» من "
                        f"{len(v)} خانة — القاعدة لا تتجاوز 5")

    # المتعامل وذو الحق بنفس الاسم — قد يكون خلطاً
    on, op_ = val("operateur", "nom_ar"), val("operateur", "prenom_ar")
    ln, lp = val("licence_modjahid", "nom_ar"), val("licence_modjahid", "prenom_ar")
    if on and ln and on == ln and op_ and lp and op_ == lp:
        warn.append("المتعامل وذو الحق بنفس الاسم واللقب — تحقّق: هل هو "
                    "Bénéféciaire يقود رخصته، أم خلطٌ بين شخصين؟")

    # رقما تعريف متبادلان بين قسمين
    for ga, gb in detect_swapped_nin(data):
        warn.append(f"🔄 رقما التعريف في «{SCHEMA[ga]['title']}» و«{SCHEMA[gb]['title']}» "
                    f"يبدوان متبادلين — كلٌّ منهما يطابق بيانات الآخر. "
                    f"تحقّق قبل أي تصحيح آخر.")

    # الخانة الثانية من رقم التعريف لا تكون إلا 0 أو 1 — تحقّق قاطع
    for grp, ninf in NIN_FIELD.items():
        nn = "".join(c for c in (val(grp, ninf) or "") if c.isdigit())
        if len(nn) == 18 and nn[1] not in "01":
            title = SCHEMA[grp]["title"]
            sx = (val(grp, "sexe") or "").upper()
            dn = val(grp, "date_naissance")
            na = val(grp, "n_acte_naissance")
            shape = ""
            try:
                y3 = f"{datetime.date.fromisoformat(dn[:10]).year % 1000:03d}" if dn else "XXX"
                a5 = f"{int(na):05d}" if na and na.isdigit() else "XXXXX"
                d2 = "1" if sx == "F" else ("0" if sx == "M" else "X")
                shape = f" — ينبغي أن يكون على صورة {nn[0]}{d2}{y3}·XXXX·{a5}·XXXX"
            except (ValueError, TypeError):
                pass
            warn.append(f"[{title}] رقم التعريف «{nn}» معطوب: خانته الثانية «{nn[1]}» "
                        f"ولا تكون إلا 0 (ذكر) أو 1 (أنثى){shape}")

    # تقاطع رقم التعريف مع بقية الحقول — يكشف خطأ في أيّهما
    for grp, ninf in NIN_FIELD.items():
        dec = decode_nin(val(grp, ninf))
        if not dec:
            continue
        title = SCHEMA[grp]["title"]
        sx = val(grp, "sexe")
        if sx and dec.get("sexe") and sx.upper() != dec["sexe"]:
            warn.append(f"[{title}] الجنس «{sx}» يخالف رقم التعريف الذي يفيد "
                        f"«{dec['sexe']}»")
        dn = val(grp, "date_naissance")
        _year_only = False
        if dn and dec.get("annee_naissance"):
            try:
                y = datetime.date.fromisoformat(dn[:10]).year
                _year_only = y != dec["annee_naissance"]
            except ValueError:
                pass
        # تصويت الشهود الثلاثة
        wit = gather_witnesses(data.get(grp) or {}, ninf)
        for key, lbl in (("annee_naissance", "سنة الميلاد"),
                         ("n_acte_naissance", "رقم شهادة الميلاد")):
            votes = wit[key]["votes"]
            win = wit[key]["winner"]
            if win is not None and len(votes) > 1:
                odd = [(v, src) for v, src in votes.items() if v != win]
                for v, src in odd:
                    warn.append(f"[{title}] {lbl}: {' و'.join(src)} يقول {v} "
                                f"بينما {' و'.join(votes[win])} يقولان {win} "
                                f"— الأرجح {win}")
            elif win is None and len(votes) > 1:
                detail = " · ".join(f"{v} ({' و'.join(src)})" for v, src in votes.items())
                warn.append(f"[{title}] {lbl}: مصادر متضاربة بلا أغلبية — {detail}")

        na = val(grp, "n_acte_naissance")
        bad_acte = bad_year = False
        if na and dec.get("n_acte_naissance"):
            try:
                bad_acte = int(na) != int(dec["n_acte_naissance"])
            except ValueError:
                pass
        if dn and dec.get("annee_naissance"):
            try:
                bad_year = (datetime.date.fromisoformat(dn[:10]).year
                            != dec["annee_naissance"])
            except ValueError:
                pass
        # تناقضان معاً ← الرقم الوطني نفسه مقروء خطأً؛ نرشد إلى صورته الصحيحة
        if bad_acte and bad_year:
            nn = "".join(c for c in val(grp, ninf) if c.isdigit())
            try:
                y3 = f"{datetime.date.fromisoformat(dn[:10]).year % 1000:03d}"
                a5 = f"{int(na):05d}"
                _sx = (val(grp, "sexe") or "").upper()
                sx = "1" if _sx == "F" else ("0" if _sx == "M"
                                             else (nn[1] if len(nn) > 1 else "X"))
                warn.append(
                    f"[{title}] رقم التعريف «{nn}» يخالف الوثيقة في السنة والشهادة معاً "
                    f"— أعد قراءته: ينبغي أن يكون على صورة "
                    f"{nn[0] if nn else 'X'}{sx}{y3}·{nn[5:9] if len(nn) >= 9 else 'XXXX'}"
                    f"·{a5}·XXXX")
            except (ValueError, TypeError):
                warn.append(f"[{title}] رقم التعريف يخالف الوثيقة في السنة والشهادة معاً")

    # ══════════════════════════════════════════════════════════════════
    #  الحقول المتوقَّعة الفارغة
    #
    #  كان الفحص يتحقّق من اتّساق ما قُرئ فقط، فيمرّ سجلٌّ نصفه فارغ
    #  بصفر تنبيهات: رخصة السياقة أمام الموظّف تحمل التاريخ والصنف
    #  والاسم، والحقول خالية، والبرنامج يقول «اجتاز الفحص».
    #
    #  لا يُدخَل تلقائياً إلا ما اجتاز الفحص بصفر تنبيهات — وهذا الوعد
    #  بلا معنى إن كان الفحص لا يرى الفراغ.
    # ══════════════════════════════════════════════════════════════════
    for grp, fields in EXPECTED.items():
        if not section_active(data, grp):
            continue
        sec = data.get(grp) or {}
        labels = {f[0]: f[1] for f in SCHEMA[grp]["fields"]}
        title = SCHEMA[grp]["title"]

        missing = [f for f in fields if str(sec.get(f) or "").strip() == ""]
        if not missing:
            continue

        # تُجمَع حسب الوثيقة التي تُقرأ منها، فينظر الموظّف في ورقة واحدة
        by_doc: dict[str, list[str]] = {}
        for f in missing:
            by_doc.setdefault(SOURCE_DOC.get(f, ""), []).append(labels.get(f, f))

        for doc, names in by_doc.items():
            where = f" — انظر في {doc}" if doc else ""
            warn.append(f"[{title}] لم تُقرأ: {' · '.join(names)}{where}")

    return warn
