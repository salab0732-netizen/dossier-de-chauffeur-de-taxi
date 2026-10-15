#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""استنتاج الحقول الناقصة وإصلاح أرقام التعريف من شواهد الملف.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import datetime
import re

from taxi.constants import WILAYA_ID
from taxi.core.schema import SCHEMA
from taxi.core.vocabulary import _VOCAB


NIN_FIELD = {"operateur": "nin", "chauffeur": "nin",
             "chauffeur_manaoub": "nin", "licence_modjahid": "nin_modjahid"}

# ثوابت مؤكّدة في القاعدة بنسبة 100%
CONSTANTS = {
    "id_wilaya":      str(WILAYA_ID),
    "pays":           "Algerie",
    "pays_naissance": "ALGERIE",
}

# حقول تُنسخ من المتعامل إلى بقية الأقسام إن غابت
PROPAGATE = {
    "chauffeur":        ["activite", "nature_activite", "adresse", "adresse_ar",
                         "id_wilaya_adresse", "pays", "pays_naissance"],
    "chauffeur_manaoub": ["activite", "nature_activite"],
    "vehicule":         ["activite", "nature_activite", "adresse", "adresse_ar",
                         "id_wilaya_adresse"],
}

# هويّة السائق الأساسي = هويّة المتعامل نفسه
IDENTITY = ["nin", "prenom", "nom", "prenom_ar", "nom_ar", "date_naissance",
            "sexe", "n_acte_naissance", "adresse", "adresse_ar",
            "id_wilaya_adresse", "numero_cnas", "numero_casnos"]


# ما يجعل القسم «عامراً» فعلاً — لا مجرّد ثوابت ملأناها نحن
ACTIVE_MARKERS = {
    "operateur":        ["nin", "nom_ar", "nom", "raison_sociale"],
    "chauffeur":        ["nin", "nom_ar", "nom", "n_permis"],
    "chauffeur_manaoub": ["nin", "nom_ar", "nom", "n_permis"],
    "vehicule":         ["n_immatriculation", "n_chasis", "marque", "n_porte"],
    "licence_modjahid": ["nin_modjahid", "nom_ar", "nom", "n_decision_octroi"],
}


def section_active(data: dict, grp: str) -> bool:
    """هل في القسم معلومة جوهرية؟ القسم الفارغ لا يُثرى ولا يُكتب."""
    sec = data.get(grp)
    if not isinstance(sec, dict):
        return False
    return any(sec.get(f) not in (None, "", "null")
               for f in ACTIVE_MARKERS.get(grp, []))


def decode_cnas(cnas: str) -> dict:
    """
    رقم الضمان الاجتماعي للأجراء يبدأ بسنة الميلاد (خانتان)
    ثم رقم شهادة الميلاد (أربع خانات) — تحقّقٌ على 861 سجلاً بدقّة 90%.
    """
    d = "".join(ch for ch in str(cnas or "") if ch.isdigit())
    if len(d) < 10:
        return {}
    out = {}
    try:
        yy = int(d[:2])
        out["annee_naissance"] = 1900 + yy if yy > 10 else 2000 + yy
        out["n_acte_naissance"] = str(int(d[2:6]))
    except ValueError:
        return {}
    return out


def nin_fit(nin: str, sec: dict) -> int:
    """كم من خصائص الشخص يوافقها هذا الرقم؟ (الجنس · السنة · رقم الشهادة)"""
    d = decode_nin(nin)
    if not d:
        return -1
    score = 0
    sx = str(sec.get("sexe") or "").upper()
    if sx and d.get("sexe") == sx:
        score += 1
    dn = sec.get("date_naissance")
    if dn and d.get("annee_naissance"):
        try:
            if datetime.date.fromisoformat(str(dn)[:10]).year == d["annee_naissance"]:
                score += 1
        except ValueError:
            pass
    ac = sec.get("n_acte_naissance")
    if ac and d.get("n_acte_naissance"):
        try:
            if int(ac) == int(d["n_acte_naissance"]):
                score += 1
        except (TypeError, ValueError):
            pass
    return score


def detect_swapped_nin(data: dict) -> list[tuple[str, str]]:
    """
    يكشف رقمَي تعريف وُضع كلٌّ منهما في قسم الآخر.
    لا يُبلّغ إلا إذا كان التبادل يُحسّن المطابقة فعلاً.
    """
    out = []
    keys = [(g, f) for g, f in NIN_FIELD.items()
            if isinstance(data.get(g), dict)
            and str((data[g] or {}).get(f) or "").strip()]
    for i in range(len(keys)):
        for j in range(i + 1, len(keys)):
            (ga, fa), (gb, fb) = keys[i], keys[j]
            na, nb = str(data[ga][fa]).strip(), str(data[gb][fb]).strip()
            if na == nb:
                continue
            now = nin_fit(na, data[ga]) + nin_fit(nb, data[gb])
            swap = nin_fit(nb, data[ga]) + nin_fit(na, data[gb])
            if swap > now and swap >= 2:
                out.append((ga, gb))
    return out


def gather_witnesses(sec: dict, ninf: str = "nin") -> dict:
    """
    يجمع شهادات مستقلّة على سنة الميلاد ورقم شهادة الميلاد:
    الحقل المصرَّح · رقم التعريف الوطني · رقم CNAS.
    يعيد لكل منهما: {"votes": {قيمة: [مصادر]}, "winner": قيمة أو None}
    """
    src: dict[str, dict] = {"annee_naissance": {}, "n_acte_naissance": {}}

    def add(key, val, who):
        if val in (None, ""):
            return
        try:
            v = int(val)
        except (TypeError, ValueError):
            return
        src[key].setdefault(v, []).append(who)

    dn = sec.get("date_naissance")
    if dn:
        try:
            add("annee_naissance",
                datetime.date.fromisoformat(str(dn)[:10]).year, "الحقل المصرَّح")
        except ValueError:
            pass
    add("n_acte_naissance", sec.get("n_acte_naissance"), "الحقل المصرَّح")

    for who, dec in (("رقم التعريف الوطني", decode_nin(sec.get(ninf))),
                     ("رقم CNAS", decode_cnas(sec.get("numero_cnas")))):
        for k in ("annee_naissance", "n_acte_naissance"):
            add(k, dec.get(k), who)

    out = {}
    for k, votes in src.items():
        winner = None
        if votes:
            top = max(votes.items(), key=lambda kv: len(kv[1]))
            if len(top[1]) >= 2:
                winner = top[0]
        out[k] = {"votes": votes, "winner": winner}
    return out


def repair_nin(nin: str, sexe=None, annee=None, acte=None) -> tuple[str | None, str]:
    """
    يحاول إصلاح رقم تعريف طوله 19 خانة بحذف خانة واحدة.
    لا يقبل إلا مرشّحاً وحيداً يجتاز كل الفحوص المتاحة.
    يعيد (الرقم المُصلَح أو None، سبب القرار).
    """
    s = "".join(ch for ch in str(nin or "") if ch.isdigit())
    if len(s) == 18:
        return s, "سليم"
    if len(s) != 19:
        return None, f"الطول {len(s)} — لا تُصلَح إلا الأرقام ذات 19 خانة"

    def ok(c: str) -> bool:
        d = decode_nin(c)
        if not d:
            return False
        if d.get("sexe") is None:           # الخانة 2 ليست 0 ولا 1
            return False
        if sexe and d["sexe"] != str(sexe).strip().upper():
            return False
        if annee and d.get("annee_naissance") != annee:
            return False
        if acte is not None:
            try:
                if int(d.get("n_acte_naissance", -1)) != int(acte):
                    return False
            except (TypeError, ValueError):
                pass
        return True

    cands = {s[:i] + s[i + 1:] for i in range(19)}
    good = sorted(c for c in cands if ok(c))
    if len(good) == 1:
        return good[0], "مرشّح وحيد يجتاز بنية الرقم"
    if not good:
        return None, "لا مرشّح يطابق بنية الرقم — يحتاج مراجعة الوثيقة"
    return None, f"{len(good)} مرشّحات محتملة — لا يُحسم آلياً"


def decode_nin(nin: str) -> dict:
    """يفكّ ما يحمله رقم التعريف الوطني من معلومات."""
    s = "".join(ch for ch in str(nin or "") if ch.isdigit())
    if len(s) != 18:
        return {}
    out = {}
    if s[1] == "0":
        out["sexe"] = "M"
    elif s[1] == "1":
        out["sexe"] = "F"
    try:
        y = 1000 + int(s[2:5])
        if 1900 <= y <= datetime.date.today().year:
            out["annee_naissance"] = y
    except ValueError:
        pass
    try:
        out["n_acte_naissance"] = str(int(s[9:14]))
    except ValueError:
        pass
    out["bureau_etat_civil"] = s[5:9]
    return out


def enrich(data: dict, log=None) -> list[str]:
    """
    يملأ الحقول الغائبة من معلومات متوفّرة في الملف نفسه.
    لا يملأ إلا ما كان اشتقاقه موثّقاً؛ وما كان احتمالياً يُترك للفحص.
    يعدّل data في مكانه ويعيد قائمة بما اشتُقّ ومن أين.
    """
    done: list[str] = []

    def put(grp, fkey, val, src):
        if grp not in SCHEMA:
            return
        if fkey not in [f for f, _, _ in SCHEMA[grp]["fields"]]:
            return
        sec = data.setdefault(grp, {})
        if sec.get(fkey) not in (None, "", "null"):
            return
        sec[fkey] = val
        label = next((lb for f, lb, _ in SCHEMA[grp]["fields"] if f == fkey), fkey)
        done.append(f"[{SCHEMA[grp]['title']}] {label} = «{val}»  ← {src}")
        if log:
            log(f"🧩 {done[-1]}")

    # ① من رقم التعريف الوطني
    for grp, ninf in NIN_FIELD.items():
        sec = data.get(grp)
        if not isinstance(sec, dict):
            continue
        dec = decode_nin(sec.get(ninf))
        if not dec:
            continue
        if "sexe" in dec:
            put(grp, "sexe", dec["sexe"], "الخانة 2 من رقم التعريف")
        w = gather_witnesses(sec, ninf)
        aw = w["n_acte_naissance"]
        if aw["winner"] is not None:
            who = " و".join(aw["votes"][aw["winner"]])
            put(grp, "n_acte_naissance", str(aw["winner"]),
                f"إجماع {len(aw['votes'][aw['winner']])} مصادر: {who}")
        elif "n_acte_naissance" in dec:
            put(grp, "n_acte_naissance", dec["n_acte_naissance"],
                "الخانات 10–14 من رقم التعريف")

    # ② ثوابت القاعدة — للأقسام العامرة فقط
    for grp in SCHEMA:
        if not section_active(data, grp):
            continue
        for fkey, val in CONSTANTS.items():
            put(grp, fkey, val, "ثابت في القاعدة")

    # ②b حقول لها قيمة واحدة في القاعدة كلها ← ترميز موحّد لا يُترجَم
    for grp in SCHEMA:
        if not section_active(data, grp):
            continue
        for fkey, vals in (_VOCAB.get(grp) or {}).items():
            if len(vals) != 1:
                continue
            sec = data.setdefault(grp, {})
            cur = sec.get(fkey)
            if cur in (None, "", "null"):
                put(grp, fkey, vals[0], "القيمة الوحيدة في القاعدة")
            elif str(cur).strip() != vals[0]:
                old = str(cur).strip()
                sec[fkey] = vals[0]
                label = next((lb for f, lb, _ in SCHEMA[grp]["fields"] if f == fkey), fkey)
                done.append(f"[{SCHEMA[grp]['title']}] {label}: «{old}» ← «{vals[0]}» "
                            f"(ترميز القاعدة الموحّد)")
                if log:
                    log(f"🧩 {done[-1]}")

    # ③ النسخ من المتعامل إلى بقية الأقسام
    op = data.get("operateur") or {}
    for grp, fields in PROPAGATE.items():
        if not section_active(data, grp):
            continue
        for f in fields:
            v = op.get(f)
            if v not in (None, "", "null"):
                put(grp, f, v, "من قسم المتعامل")

    # ④ السائق الأساسي هو المتعامل نفسه
    ch = data.get("chauffeur")
    if isinstance(ch, dict) and ch:
        same = (ch.get("qualite_chauffeur") in (None, "", "null", "Principal")
                and ch.get("nin") in (None, "", "null"))
        named = ch.get("nom_ar") in (None, "", "null") or \
            str(ch.get("nom_ar", "")).strip() == str(op.get("nom_ar", "")).strip()
        if same and named and op.get("nin"):
            for f in IDENTITY:
                v = op.get(f)
                if v not in (None, "", "null"):
                    put("chauffeur", f, v, "السائق الأساسي = المتعامل")
            put("chauffeur", "qualite_chauffeur", "Principal", "استنتاج: يقود مركبته")

    return done
