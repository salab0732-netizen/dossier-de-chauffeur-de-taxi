#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
═══════════════════════════════════════════════════════════════════════
  نظام إدخال بيانات سيارات الأجرة  —  ولاية البيض (32)
  Taxi Licensing Data Entry System                         v2.0
═══════════════════════════════════════════════════════════════════════
  المزوّدون المدعومون :  Google Gemini  (مجاني)  |  Anthropic Claude
  المتطلبات           :  pip install google-genai pymupdf openpyxl
═══════════════════════════════════════════════════════════════════════
"""

from __future__ import annotations

import base64
import datetime
import json
import os
import random
import shutil
import sys
import threading
import time
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

APP_NAME    = "نظام إدخال بيانات سيارات الأجرة"
APP_VERSION = "5.7.1"
WILAYA_ID   = 32
CONFIG_FILE = Path(__file__).resolve().parent / "config.json"

DEFAULT_EXCEL = r"F:\Documents\تحديث ملف سيارات الاجرة\TAXI 32000.xlsx"

GEMINI_MODELS = [
    "gemini-3.6-flash",
    "gemini-3.6-pro",
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-2.0-flash",
]
CLAUDE_MODELS = [
    "claude-sonnet-4-5",
    "claude-opus-4-5",
    "claude-haiku-4-5",
]

# ══════════════════════════════════════════════════════════════════════
#  مخطط البيانات  —  يقود الاستخراج والمراجعة والكتابة في Excel
#  كل حقل :  (المفتاح في Excel ، التسمية العربية ، وصف للنموذج)
# ══════════════════════════════════════════════════════════════════════

SCHEMA = {
    "operateur": {
        "sheet": "Operateur",
        "title": "المتعامل (السائق المالك)",
        "icon":  "🏠",
        "key":   "n-porte",
        "porte": "n-porte",
        "dup":   [("n-porte",), ("nin",)],
        "fields": [
            ("id_wilaya",             "رقم الولاية",              "دائماً 32"),
            ("nin",                   "رقم التعريف الوطني",       "18 رقم"),
            ("prenom",                "الاسم (FR)",               "بالحروف اللاتينية"),
            ("nom",                   "اللقب (FR)",               "بالحروف اللاتينية"),
            ("prenom_ar",             "الاسم (AR)",               "بالعربية"),
            ("nom_ar",                "اللقب (AR)",               "بالعربية"),
            ("date_naissance",        "تاريخ الميلاد",            "YYYY-MM-DD"),
            ("sexe",                  "الجنس",                    "M أو F"),
            ("n_acte_naissance",      "رقم شهادة الميلاد",        ""),
            ("adresse",               "العنوان (FR)",             ""),
            ("adresse_ar",            "العنوان (AR)",             ""),
            ("id_wilaya_adresse",     "ولاية العنوان",            "دائماً 32"),
            ("n-porte",               "رقم الباب",                "رقم الرخصة على باب التاكسي"),
            ("date_debut_contrat",    "بداية الرخصة",             "YYYY-MM-DD"),
            ("date_fin_contrat",      "نهاية الرخصة",             "YYYY-MM-DD"),
            ("qualite_oper_physique", "صفة المتعامل",             "Locatiare إن كان يستأجر الرخصة، Bénéféciaire إن كان هو ذا الحق"),
            ("activite",              "النشاط",                   "Transports par Taxi"),
            ("nature_activite",       "طبيعة النشاط",             "Individuelle أو inter-wilayas"),
            ("numero_casnos",         "رقم CASNOS",               ""),
            ("numero_cnas",           "رقم CNAS",                 ""),
            ("prenom_pere",           "اسم الأب",                 "من البطاقة الرمادية أو شهادة الميلاد"),
            ("prenom_pere_ar",        "اسم الأب (AR)",            ""),
            ("prenom_mere",           "اسم الأم",                 ""),
            ("prenom_mere_ar",        "اسم الأم (AR)",            ""),
            ("nom_mere",              "لقب الأم",                 ""),
            ("nom_mere_ar",           "لقب الأم (AR)",            ""),
            ("numero_telephone",      "رقم الهاتف",               ""),
            ("n_dossier",             "رقم قيد المتعامل",         ""),
            ("date_debut_activite",   "تاريخ بداية النشاط",       "YYYY-MM-DD"),
            ("pays_naissance",        "بلد الميلاد",              "Algérie عادةً"),
            ("pays",                  "البلد",                    "Algérie عادةً"),
            # ── حقول الشركة: تُملأ فقط إن كان المتعامل شخصاً معنوياً ──
            ("type_personne_morale",  "شخص معنوي",                "oui إن كان شركة، non إن كان شخصاً"),
            ("raison_sociale",        "اسم الشركة",               "مثل EURL DJEBBAR"),
            ("raison_sociale_ar",     "اسم الشركة (AR)",          ""),
            ("sigle",                 "اختصار اسم الشركة",        ""),
            ("form_entreprise",       "شكل الشركة",               "EURL / SARL / SPA…"),
            ("secteur_juridique",     "القطاع القانوني",          "Prive / Public / Mixte"),
            ("NRC",                   "رقم السجل التجاري",        ""),
            ("date_registre_commerce","تاريخ السجل التجاري",      "YYYY-MM-DD"),
            ("NIF",                   "رقم التعريف الضريبي",      ""),
            ("adresse_sociale",       "العنوان الاجتماعي",        ""),
            ("adresse_sociale_ar",    "العنوان الاجتماعي (AR)",   ""),
            ("nom_gerant",            "اسم ولقب المسيّر",          ""),
            ("nom_gerant_ar",         "اسم ولقب المسيّر (AR)",     ""),
            ("num_agrement",          "رقم الاعتماد",             ""),
            ("date_agrement",         "تاريخ الاعتماد",           "YYYY-MM-DD"),
            ("cessation_activite",    "متوقف عن النشاط",          "oui / non"),
        ],
    },
    "chauffeur": {
        "sheet": "Chauffeur",
        "title": "السائق",
        "icon":  "🚗",
        "key":   "nin",          # السائق يُميَّز برقم تعريفه الوطني لا برقم الباب
        "porte": None,           # لا عمود لرقم الباب في هذه الورقة
        # إن غاب NIN نسقط إلى اللقب+الاسم+تاريخ الميلاد
        "dup":   [("nin",), ("nom_ar", "prenom_ar", "date_naissance")],
        "fields": [
            ("id_wilaya",                 "رقم الولاية",          "دائماً 32"),
            ("activite",                  "النشاط",               "Transports par Taxi"),
            ("nature_activite",           "طبيعة النشاط",         "Individuelle أو inter-wilayas"),
            ("qualite_chauffeur",         "الصفة",                "Principal للسائق الأساسي، Doubleur للمناوب"),
            ("nin",                       "رقم التعريف الوطني",   "18 رقم"),
            ("prenom",                    "الاسم (FR)",           ""),
            ("nom",                       "اللقب (FR)",           ""),
            ("prenom_ar",                 "الاسم (AR)",           ""),
            ("nom_ar",                    "اللقب (AR)",           ""),
            ("date_naissance",            "تاريخ الميلاد",        "YYYY-MM-DD"),
            ("sexe",                      "الجنس",                "M أو F"),
            ("adresse",                   "العنوان (FR)",         ""),
            ("adresse_ar",                "العنوان (AR)",         ""),
            ("id_wilaya_adresse",         "ولاية العنوان",        "دائماً 32"),
            ("numero_cnas",               "رقم CNAS",             ""),
            ("numero_casnos",             "رقم CASNOS",           ""),
            ("n_permis",                  "رقم رخصة السياقة",     ""),
            ("date_delivrance_permis",    "تاريخ رخصة السياقة",   "YYYY-MM-DD"),
            ("categorie_permis",          "صنف الرخصة",           "B أو D"),
            ("groupage",                  "فصيلة الدم",           ""),
            ("n_certificat_formation",    "رقم شهادة التكوين",    ""),
            ("date_certificat_formation", "تاريخ شهادة التكوين",  "YYYY-MM-DD"),
            ("n_telephone",               "رقم الهاتف",           ""),
            ("n_acte_naissance",          "رقم شهادة الميلاد",    ""),
            ("n_dossier",                 "رقم قيد المتعامل",     ""),
            ("n_delivrance_licence_exploitation",  "رقم إصدار رخصة الاستغلال", ""),
            ("date_delivrance_licence_exploitation", "تاريخ إصدار رخصة الاستغلال", "YYYY-MM-DD"),
            ("n_registre_siege",          "رقم دفتر المقاعد",     ""),
            ("date_registre_siege",       "تاريخ دفتر المقاعد",   "YYYY-MM-DD"),
            ("validite_categorie_permis", "صلاحية صنف الرخصة",    "YYYY-MM-DD"),
            ("date_fin_control_medical",  "نهاية الفحص الطبي",    "YYYY-MM-DD"),
            ("date_mise_a_jour_securite_sociale", "تحيين الضمان الاجتماعي", "YYYY-MM-DD"),
            ("pays_naissance",            "بلد الميلاد",          "Algérie عادةً"),
            ("pays",                      "البلد",                "Algérie عادةً"),
        ],
    },
    "chauffeur_manaoub": {
        "sheet": "Chauffeur",
        "title": "السائق المناوب",
        "icon":  "🔁",
        "key":   "nin",          # السائق يُميَّز برقم تعريفه الوطني لا برقم الباب
        "porte": None,           # لا عمود لرقم الباب في هذه الورقة
        # إن غاب NIN نسقط إلى اللقب+الاسم+تاريخ الميلاد
        "dup":   [("nin",), ("nom_ar", "prenom_ar", "date_naissance")],
        "fields": [
            ("id_wilaya",                 "رقم الولاية",          "دائماً 32"),
            ("activite",                  "النشاط",               "Transports par Taxi"),
            ("nature_activite",           "طبيعة النشاط",         "Individuelle أو inter-wilayas"),
            ("qualite_chauffeur",         "الصفة",                "Principal للسائق الأساسي، Doubleur للمناوب"),
            ("nin",                       "رقم التعريف الوطني",   "18 رقم"),
            ("prenom",                    "الاسم (FR)",           ""),
            ("nom",                       "اللقب (FR)",           ""),
            ("prenom_ar",                 "الاسم (AR)",           ""),
            ("nom_ar",                    "اللقب (AR)",           ""),
            ("date_naissance",            "تاريخ الميلاد",        "YYYY-MM-DD"),
            ("sexe",                      "الجنس",                "M أو F"),
            ("adresse",                   "العنوان (FR)",         ""),
            ("adresse_ar",                "العنوان (AR)",         ""),
            ("id_wilaya_adresse",         "ولاية العنوان",        "دائماً 32"),
            ("numero_cnas",               "رقم CNAS",             ""),
            ("numero_casnos",             "رقم CASNOS",           ""),
            ("n_permis",                  "رقم رخصة السياقة",     ""),
            ("date_delivrance_permis",    "تاريخ رخصة السياقة",   "YYYY-MM-DD"),
            ("categorie_permis",          "صنف الرخصة",           "B أو D"),
            ("groupage",                  "فصيلة الدم",           ""),
            ("n_certificat_formation",    "رقم شهادة التكوين",    ""),
            ("date_certificat_formation", "تاريخ شهادة التكوين",  "YYYY-MM-DD"),
            ("n_telephone",               "رقم الهاتف",           ""),
            ("n_acte_naissance",          "رقم شهادة الميلاد",    ""),
            ("n_dossier",                 "رقم قيد المتعامل",     ""),
            ("n_delivrance_licence_exploitation",  "رقم إصدار رخصة الاستغلال", ""),
            ("date_delivrance_licence_exploitation", "تاريخ إصدار رخصة الاستغلال", "YYYY-MM-DD"),
            ("n_registre_siege",          "رقم دفتر المقاعد",     ""),
            ("date_registre_siege",       "تاريخ دفتر المقاعد",   "YYYY-MM-DD"),
            ("validite_categorie_permis", "صلاحية صنف الرخصة",    "YYYY-MM-DD"),
            ("date_fin_control_medical",  "نهاية الفحص الطبي",    "YYYY-MM-DD"),
            ("date_mise_a_jour_securite_sociale", "تحيين الضمان الاجتماعي", "YYYY-MM-DD"),
            ("pays_naissance",            "بلد الميلاد",          "Algérie عادةً"),
            ("pays",                      "البلد",                "Algérie عادةً"),
        ],
    },
    "vehicule": {
        "sheet": "Véhicule",
        "title": "المركبة",
        "icon":  "🚙",
        "key":   "n_porte",
        "porte": "n_porte",
        "dup":   [("n_porte",), ("n_immatriculation",)],
        "fields": [
            ("id_wilaya",                    "رقم الولاية",           "دائماً 32"),
            ("activite",                     "النشاط",                "Transports par Taxi"),
            ("nature_activite",              "طبيعة النشاط",          ""),
            ("n_porte",                      "رقم الباب",             ""),
            ("n_certificat_cerculation",     "رقم رخصة السير",        ""),
            ("date_fin_validite_certificat", "نهاية صلاحية الرخصة",   "YYYY-MM-DD"),
            ("n_carte_gris",                 "رقم البطاقة الرمادية",  ""),
            ("adresse",                      "العنوان (FR)",          ""),
            ("adresse_ar",                   "العنوان (AR)",          ""),
            ("id_wilaya_adresse",            "ولاية العنوان",         "دائماً 32"),
            ("n_immatriculation",            "رقم التسجيل",           "مثال 05427.112.32"),
            ("n_chasis",                     "الرقم التسلسلي VIN",    ""),
            ("type",                         "الطراز TYPE",           ""),
            ("marque",                       "الصنف MARQUE",          "RENAULT / TOYOTA ..."),
            ("genre",                        "النوع GENRE",           ""),
            ("carrossere",                   "الهيكل",                ""),
            ("n_siege",                      "عدد المقاعد",           ""),
            ("puissance",                    "القوة",                 ""),
            ("energie",                      "الطاقة",                "ESSENCE / GASOIL / GPL"),
            ("annee_premier_utilisation",    "سنة أول استعمال",       ""),
            ("num_dossier",                  "رقم قيد المتعامل",      ""),
            ("num_ligne",                    "رقم الخط",              ""),
            ("date_debut_activite_vehecule", "تاريخ بداية نشاط المركبة", "YYYY-MM-DD"),
            ("charge utile",                 "الحمولة المقيدة",       "من البطاقة الرمادية"),
            ("poids total",                  "جملة الحمولة",          "من البطاقة الرمادية"),
        ],
    },
    "licence_modjahid": {
        "sheet": "Licence Modjahid",
        "title": "رخصة المجاهد (ذو الحق)",
        "icon":  "📜",
        "key":   "n_porte",
        "porte": "n_porte",
        "dup":   [("n_porte",), ("nin_modjahid",)],
        "fields": [
            ("id_wilaya",            "رقم الولاية",              "دائماً 32"),
            ("n_porte",              "رقم الباب",                ""),
            ("nin_modjahid",         "NIN المجاهد",              ""),
            ("prenom",               "اسم المجاهد (FR)",         ""),
            ("nom",                  "لقب المجاهد (FR)",         ""),
            ("prenom_ar",            "اسم المجاهد (AR)",         ""),
            ("nom_ar",               "لقب المجاهد (AR)",         ""),
            ("date_naissance",       "تاريخ ميلاد المجاهد",      "YYYY-MM-DD"),
            ("n_acte_naissance",     "رقم شهادة الميلاد",        ""),
            ("adresse",              "عنوان المجاهد (FR)",       ""),
            ("adresse_ar",           "عنوان المجاهد (AR)",       ""),
            ("type_licence",         "نوع الرخصة",               "Délivrée أو Transférée"),
            ("n_decision_octroi",    "رقم قرار المنح",           ""),
            ("date_decision_octroi", "تاريخ قرار المنح",         "YYYY-MM-DD"),
            ("concentration_exploitation_licence_modjahid", "مكان الاستغلال",
             "البلدية المذكورة في القرار الولائي"),
        ],
    },
}


# ورقة «Referances» هي المرجعية الرسمية للمصطلحات.
# المفتاح = عنوان العمود فيها (الصف 1)، والقيمة = الحقول التي يحكمها.
REF_SHEET = "Referances"
REF_MAP = {
    "Qualité Operateur Taxi": [("operateur", "qualite_oper_physique")],
    "Qualite chauffeur":      [("chauffeur", "qualite_chauffeur"),
                               ("chauffeur_manaoub", "qualite_chauffeur")],
    "Type Licence":           [("licence_modjahid", "type_licence")],
    "Form Entreprise":        [("operateur", "form_entreprise")],
    "Secteur Juridique":      [("operateur", "secteur_juridique")],
    "oui / non":              [("operateur", "type_personne_morale"),
                               ("operateur", "cessation_activite")],
    "Energie":                [("vehicule", "energie")],
    "Marque":                 [("vehicule", "marque")],
    "Groupage":               [("chauffeur", "groupage"),
                               ("chauffeur_manaoub", "groupage")],
    "Sexe":                   [("operateur", "sexe"), ("chauffeur", "sexe"),
                               ("chauffeur_manaoub", "sexe"),
                               ("licence_modjahid", "sexe")],
    "Transports par Taxi":    [("operateur", "nature_activite"),
                               ("chauffeur", "nature_activite"),
                               ("chauffeur_manaoub", "nature_activite"),
                               ("vehicule", "nature_activite")],
}
# حقول قيمتها هي عنوان العمود نفسه
REF_HEADER_AS_VALUE = {
    "Transports par Taxi": [("operateur", "activite"), ("chauffeur", "activite"),
                            ("chauffeur_manaoub", "activite"), ("vehicule", "activite")],
}


def read_reference_sheet(excel_path: str) -> dict[str, list[str]]:
    """يقرأ ورقة Referances: {عنوان العمود: [القيم المسموحة]}."""
    import openpyxl
    out: dict[str, list[str]] = {}
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    except Exception:
        return out
    try:
        if REF_SHEET not in wb.sheetnames:
            return out
        rows = list(wb[REF_SHEET].iter_rows(values_only=True))
        if len(rows) < 2:
            return out
        ncols = max(len(r) for r in rows)
        for ci in range(ncols):
            head = rows[0][ci] if ci < len(rows[0]) else None
            if head is None or str(head).strip() in ("", "_"):
                continue
            vals: list[str] = []
            # الصف 2 قد يحمل أول قيمة مسموحة
            v2 = rows[1][ci] if ci < len(rows[1]) else None
            if v2 is not None and str(v2).strip() not in ("", "_"):
                vals.append(str(v2).strip())
            for r in rows[2:]:
                v = r[ci] if ci < len(r) else None
                if v is not None and str(v).strip() not in ("", "_"):
                    s = str(v).strip()
                    if s not in vals:
                        vals.append(s)
            if vals:
                out[str(head).strip()] = vals
    finally:
        wb.close()
    return out


# حقول ذات قيم محصورة — تُتعلَّم من الملف نفسه لا تُفترض
ENUM_FIELDS = {
    ("operateur", "qualite_oper_physique"), ("operateur", "nature_activite"),
    ("operateur", "activite"),
    ("chauffeur", "qualite_chauffeur"), ("chauffeur", "nature_activite"),
    ("chauffeur", "activite"),
    ("chauffeur_manaoub", "qualite_chauffeur"), ("chauffeur_manaoub", "nature_activite"),
    ("chauffeur_manaoub", "activite"),
    ("vehicule", "nature_activite"), ("vehicule", "activite"),
    ("vehicule", "energie"), ("vehicule", "genre"), ("vehicule", "carrossere"),
    ("licence_modjahid", "type_licence"),
    ("operateur", "pays"), ("operateur", "pays_naissance"),
    ("chauffeur", "pays"), ("chauffeur", "pays_naissance"),
    ("chauffeur_manaoub", "pays"), ("chauffeur_manaoub", "pays_naissance"),
    ("vehicule", "marque"),
    ("chauffeur", "groupage"), ("chauffeur_manaoub", "groupage"),
    ("operateur", "sexe"), ("chauffeur", "sexe"),
    ("operateur", "form_entreprise"), ("operateur", "secteur_juridique"),
    ("operateur", "type_personne_morale"), ("operateur", "cessation_activite"),
    ("chauffeur_manaoub", "sexe"),
}
_VOCAB: dict = {}


def learn_vocabulary(excel_path: str, max_vals: int = 12) -> dict:
    """
    مصدران للمصطلحات، بهذا الترتيب:
      ① ورقة Referances — المرجعية الرسمية (لها الأولوية المطلقة)
      ② القيم المستعملة فعلاً في أوراق البيانات — لما لا تغطّيه المرجعية
    """
    import openpyxl
    from collections import Counter
    vocab: dict = {}
    valid = {(g, f) for g, grp in SCHEMA.items() for f, _, _ in grp["fields"]}

    # ① المرجعية الرسمية
    ref = read_reference_sheet(excel_path)
    for head, targets in REF_MAP.items():
        if head in ref:
            for g, f in targets:
                if (g, f) in valid:
                    vocab.setdefault(g, {})[f] = list(ref[head])
    for head, targets in REF_HEADER_AS_VALUE.items():
        if head in ref:
            for g, f in targets:
                if (g, f) in valid:
                    vocab.setdefault(g, {})[f] = [head]
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    except Exception:
        return vocab
    try:
        for gkey, grp in SCHEMA.items():
            sheet = grp["sheet"]
            if sheet not in wb.sheetnames:
                continue
            ws = wb[sheet]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 3:
                continue
            head = {str(c).strip(): i for i, c in enumerate(rows[0]) if c is not None}
            for fkey, _, _ in grp["fields"]:
                if (gkey, fkey) not in ENUM_FIELDS:
                    continue
                if vocab.get(gkey, {}).get(fkey):
                    continue          # المرجعية الرسمية تسبق الاستعمال
                ci = None
                for k, i in head.items():
                    if _norm(k) == _norm(fkey):
                        ci = i
                        break
                if ci is None:
                    continue
                cnt = Counter()
                for r in rows[2:]:
                    if ci < len(r) and r[ci] not in (None, ""):
                        cnt[str(r[ci]).strip()] += 1
                # نتجاهل القيم النادرة جداً (غالباً أخطاء إدخال سابقة)
                total = sum(cnt.values()) or 1
                keep = [v for v, c in cnt.most_common(max_vals) if c / total >= 0.004]
                if keep:
                    vocab.setdefault(gkey, {})[fkey] = keep
    finally:
        wb.close()
    return vocab


def build_prompt() -> str:
    """تعليمات الاستخراج — بدون مثال JSON (المخطط المُلزِم يتكفّل بالبنية)."""
    lines = [
        "أنت خبير في قراءة وثائق رخص استغلال سيارات الأجرة الجزائرية (ولاية البيض، رقم 32).",
        "اقرأ كل الصور المرفقة — وهي وثائق ممسوحة ضوئياً — واستخرج منها المعلومات المطلوبة.",
        "",
        "قواعد صارمة:",
        "• الأرقام بالأرقام اللاتينية (1290 وليس ١٢٩٠).",
        "• التواريخ بصيغة YYYY-MM-DD حصراً.",
        "• أي معلومة لم تجدها في الوثائق: اتركها فارغة — لا تخمّن أبداً.",
        "• id_wilaya دائماً «32».",
        "",
        "═══ قاعدة حاسمة: من هو صاحب كل اسم ═══",
        "الملف يحوي شخصين مختلفين على الأقل. الخلط بينهما أخطر خطأ ممكن.",
        "الوثائق نفسها هي الدليل — طابِق الأسماء بينها:",
        "",
        "① المتعامل (operateur) = سائق سيارة الأجرة **مالك المركبة**.",
        "   اسمه هو المكتوب على **البطاقة الرمادية** و**رخصة السياقة** — وهما متطابقان.",
        "   هو من يشغّل التاكسي فعلياً. غالباً رجل.",
        "   خُذ منه: اللقب، الاسم، رقم تعريفه الوطني، تاريخ ميلاده، عنوانه، CNAS/CASNOS.",
        "   ⚠️ هذا القسم ليس لصاحب الحق ولا للمجاهد — بل للسائق المالك.",
        "",
        "② رخصة المجاهد (licence_modjahid) = **ذو الحق** الذي مُنحت الرخصة باسمه.",
        "   اسمه هو المكتوب على **القرار الولائي** و**بطاقة التعريف الوطنية** المرفقة له",
        "   — وهما متطابقان. غالباً امرأة (أرملة أو ابنة مجاهد/شهيد)، ويختلف عن ①.",
        "   خُذ منه: اللقب، الاسم، رقم تعريفه الوطني، تاريخ ميلاده،",
        "   ورقم قرار المنح وتاريخه من القرار الولائي.",
        "",
        "③ السائق (chauffeur) — وله ثلاث صفات، اختر الصحيحة بحسب الوثائق:",
        "   • Principal = مالك المركبة نفسه، أي الشخص ① أعلاه. هذه الحالة الغالبة:",
        "     انسخ بياناته من ① وأضف رخصة سياقته وصنفها وشهادة التكوين وفصيلة دمه.",
        "   • Doubleur = السائق المناوب (الإضافي). يميّزه وجود **رخصة سائق مناوب**",
        "     إلى جانب رخصة سياقته. ضعه في قسم chauffeur_manaoub لا هنا.",
        "   • Chauffeur societe Taxi = سائق أجير لدى شركة سيارات أجرة.",
        "     يميّزه أن المتعامل شركة (EURL أو SARL…) لا شخصاً طبيعياً،",
        "     وتجد في الوثائق عقد عمل أو انتساب للشركة.",
        "",
        "④ السائق المناوب (chauffeur_manaoub) — إن وُجد فقط، بصفة Doubleur.",
        "   له **رخصة سياقة** و**رخصة سائق مناوب/إضافي**. شخص مستقلّ تماماً.",
        "   إن لم تجد رخصة سائق مناوب في الوثائق فاترك هذا القسم فارغاً كلياً.",
        "",
        "⑤ إن كان المتعامل **شركة** لا شخصاً طبيعياً:",
        "   املأ في قسم operateur حقول الشركة (اسمها، شكلها، سجلّها التجاري،",
        "   عنوانها الاجتماعي، اسم مسيّرها) واترك حقول الشخص الطبيعي فارغة،",
        "   واجعل طبيعة النشاط Societe Taxi.",
        "",
        "═══ رقم التعريف الوطني — اقرأه بتأنٍّ شديد ═══",
        "طوله **18 رقماً بالضبط**. عُدّ الأرقام قبل أن تكتبها، ثم عُدّها ثانية.",
        "أكثر الأخطاء شيوعاً زيادة رقم أو نقصانه.",
        "",
        "بنيته تساعدك على التحقّق:",
        "   الخانة  1      : رقم السلسلة (1 أو 2)",
        "   الخانة  2      : الجنس — 0 للذكر · 1 للأنثى",
        "   الخانات 3–5    : آخر ثلاثة أرقام من سنة الميلاد (979 لمواليد 1979)",
        "   الخانات 6–9    : رمز مكتب الحالة المدنية",
        "   الخانات 10–14  : رقم شهادة الميلاد مسبوقاً بأصفار",
        "   الخانات 15–18  : رقم تسلسلي",
        "",
        "فإن كان المولود سنة 1979 ذكراً، فرقمه يبدأ حتماً بـ «10979».",
        "قارن ما قرأته بتاريخ الميلاد المكتوب: إن لم يتطابقا فقد أخطأت القراءة.",
        "",
        "⚠️ في بطاقة التعريف رقمان مختلفان — لا تخلط بينهما:",
        "   • رقم التعريف الوطني: 18 خانة، في وسط البطاقة بجانب عبارة",
        "     «رقم التعريف الوطني». هذا هو المطلوب.",
        "   • رقم البطاقة: 9 أو 10 خانات، مطبوع في أعلى البطاقة. غير مطلوب،",
        "     ولا يوضع أبداً في خانة رقم شهادة الميلاد.",
        "   رقم شهادة الميلاد لا يتجاوز 5 خانات، وتجده في شهادة الميلاد",
        "   أو في الخانات 10–14 من رقم التعريف الوطني.",
        "",
        "═══ أمانة النقل ═══",
        "• انسخ الأسماء **حرفياً** كما هي مكتوبة في الوثيقة — لا تُصحّح إملاءً ولا تُوحّد رسماً.",
        "• لا تخترع تاريخاً أبداً. إن ظهرت السنة فقط ولم يظهر اليوم أو الشهر،",
        "  فاترك الحقل فارغاً — لا تُكمله بيوم أو شهر من عندك.",
        "• انتبه: تاريخ إصدار البطاقة وتاريخ انتهائها ليسا تاريخ الميلاد.",
        "• في البطاقة الرمادية: «القوة» (PUISSANCE) و«عدد المقاعد» (PLACES ASSISES)",
        "  حقلان متجاوران ومختلفان — لا تخلط بينهما.",
        "• الرقم التسلسلي للمركبة (n_chasis / VIN) خليطٌ من حروف وأرقام،",
        "  طوله 8 أو 17 خانة غالباً. لا يكون أبداً 18 رقماً بحتاً.",
        "  إن وجدت 18 رقماً بحتاً فذاك رقم تعريف وطني لا رقم تسلسلي.",
        "• رقم الباب (n_porte / n-porte) واحد في كل الأقسام — انسخه كما هو في كل قسم.",
        "",
        "الوثائق التي قد تجدها: رخصة الاستغلال المؤقتة · بطاقة التعريف الوطنية ·",
        "رخصة السياقة · البطاقة الرمادية · رخصة سائق مناوب · شهادة التكوين ·",
        "شهادات CNAS/CASNOS · القرار الولائي · شهادة الميلاد.",
    ]

    if _VOCAB:
        lines += ["", "═══ قيم إلزامية — استعمل هذه الصيغ حرفياً ═══",
                  "هذه هي القيم المعتمدة في قاعدة البيانات. اختر منها ولا تخترع صيغة جديدة",
                  "ولا تترجم إلى العربية، حتى لو كانت الوثيقة بالعربية:"]
        for gkey, flds in _VOCAB.items():
            title = SCHEMA[gkey]["title"]
            for fkey, vals in flds.items():
                label = next((lb for f, lb, _ in SCHEMA[gkey]["fields"] if f == fkey), fkey)
                lines.append(f"  [{title}] {label} ({fkey}): " + " | ".join(vals))
    return "\n".join(lines)


def build_response_schema(flavor: str = "gemini") -> dict:
    """
    مخطط مُلزِم يضمن بنية صالحة دائماً.
    flavor='gemini'    → صيغة OpenAPI  (nullable: true)
    flavor='anthropic' → صيغة JSON Schema  (type: ["string","null"])
    """
    anthro = flavor == "anthropic"

    def field(label: str, hint: str) -> dict:
        d = {"description": label + (f" — {hint}" if hint else "")}
        if anthro:
            d["type"] = ["string", "null"]
        else:
            d["type"] = "string"
            d["nullable"] = True
        return d

    props: dict = {}
    for gkey, grp in SCHEMA.items():
        props[gkey] = {
            "type": "object",
            "description": grp["title"],
            "properties": {f: field(lb, hn) for f, lb, hn in grp["fields"]},
        }
    props["notes"] = field("ملاحظات أو معلومات مهمة غير مصنّفة", "")
    return {"type": "object", "properties": props, "required": list(SCHEMA.keys())}


# ══════════════════════════════════════════════════════════════════════
#  الإعدادات
# ══════════════════════════════════════════════════════════════════════

def load_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_config(cfg: dict) -> bool:
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════
#  معالجة PDF
# ══════════════════════════════════════════════════════════════════════

def pdf_to_jpegs(pdf_path: str, zoom: float = 2.0, progress=None) -> list[bytes]:
    """تحويل كل صفحة PDF إلى صورة JPEG."""
    try:
        import pymupdf as fitz           # الاسم الحديث
    except ImportError:
        import fitz                      # الاسم القديم
    doc = fitz.open(pdf_path)
    out: list[bytes] = []
    total = len(doc)
    try:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            out.append(pix.tobytes("jpeg"))
            if progress:
                progress(i + 1, total)
    finally:
        doc.close()
    return out


# ══════════════════════════════════════════════════════════════════════
#  مزوّدو الذكاء الاصطناعي
# ══════════════════════════════════════════════════════════════════════

def _pkg_version(name: str) -> str:
    try:
        from importlib.metadata import version
        return version(name)
    except Exception:
        return "غير مثبّتة"


def sort_models(names) -> list[str]:
    """ترتيب النماذج: الأحدث إصداراً أولاً، والمستقرّ قبل التجريبي."""
    import re

    def key(n: str):
        ver = re.search(r"(\d+)(?:[.\-](\d+))?", n)
        major = int(ver.group(1)) if ver else 0
        minor = int(ver.group(2)) if ver and ver.group(2) else 0
        preview = any(w in n for w in ("preview", "exp", "-latest", "thinking"))
        flash = "flash" in n
        return (-major, -minor, preview, not flash, n)

    return sorted(set(names), key=key)


class ProviderError(Exception):
    pass


class Cancelled(Exception):
    pass


# أكواد تستوجب إعادة المحاولة: ازدحام، تجاوز حصّة، أعطال خوادم مؤقتة
RETRY_CODES = (429, 500, 502, 503, 504, 529)
RETRY_WORDS = ("unavailable", "overloaded", "high demand", "rate limit",
               "resource_exhausted", "deadline", "timeout", "temporarily",
               "connection reset", "connection aborted", "try again")


def is_retryable(e: Exception) -> bool:
    """هل الخطأ عابر ويستحق إعادة المحاولة؟"""
    code = getattr(e, "code", None) or getattr(e, "status_code", None)
    if isinstance(code, int) and code in RETRY_CODES:
        return True
    s = str(e).lower()
    if any(str(c) in s for c in RETRY_CODES) and "error" in s:
        return True
    return any(w in s for w in RETRY_WORDS)


def retry_call(fn, *, attempts: int = 5, base: float = 3.0,
               on_retry=None, should_stop=None):
    """تنفيذ مع إعادة محاولة وتراجع أُسّي وتشويش عشوائي."""
    for i in range(1, attempts + 1):
        if should_stop and should_stop():
            raise Cancelled()
        try:
            return fn()
        except Cancelled:
            raise
        except Exception as e:
            if i >= attempts or not is_retryable(e):
                raise
            # تشويش نسبي ±15% — يبقى التزايد الأُسّي مضموناً عند أي مقياس
            delay = min(base * (2 ** (i - 1)), 45) * random.uniform(0.85, 1.15)
            if on_retry:
                on_retry(i, attempts, delay, e)
            waited = 0.0
            while waited < delay:
                if should_stop and should_stop():
                    raise Cancelled()
                time.sleep(0.25)
                waited += 0.25


class GeminiProvider:
    name = "Gemini"

    def __init__(self, api_key: str, model: str):
        self.api_key = api_key
        self.model = model

    def check(self) -> str:
        try:
            from google import genai
        except ImportError:
            raise ProviderError("المكتبة غير مثبّتة — شغّل: pip install google-genai")
        client = genai.Client(api_key=self.api_key)
        names = []
        for m in client.models.list():
            n = getattr(m, "name", "") or ""
            if "generateContent" in (getattr(m, "supported_actions", None) or ["generateContent"]):
                names.append(n.replace("models/", ""))
        return f"الاتصال ناجح — {len(names)} نموذج متاح"

    def list_models(self) -> list[str]:
        from google import genai
        client = genai.Client(api_key=self.api_key)
        out = []
        for m in client.models.list():
            n = (getattr(m, "name", "") or "").replace("models/", "")
            if not n or any(w in n for w in ("embedding", "aqa", "imagen", "veo", "tts")):
                continue
            out.append(n)
        return sort_models(set(out))

    def extract(self, images: list[bytes], prompt: str) -> str:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=self.api_key)
        parts = [types.Part.from_bytes(data=img, mime_type="image/jpeg") for img in images]
        parts.append(types.Part.from_text(text=prompt))
        cfg = types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=build_response_schema("gemini"),   # بنية مضمونة
            temperature=0.0,
            max_output_tokens=16384,
        )
        try:
            resp = client.models.generate_content(model=self.model, contents=parts, config=cfg)
        except Exception as e:
            # بعض النماذج لا تدعم المخطط المُلزِم — نعيد المحاولة بدونه
            if "schema" not in str(e).lower() and "response_schema" not in str(e).lower():
                raise
            resp = client.models.generate_content(
                model=self.model, contents=parts,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0, max_output_tokens=16384),
            )
        txt = (resp.text or "").strip()
        if not txt:
            reason = "؟"
            try:
                reason = str(resp.candidates[0].finish_reason)
            except Exception:
                pass
            raise ProviderError(
                f"أعاد النموذج رداً فارغاً (سبب التوقّف: {reason}). "
                "جرّب نموذجاً آخر أو قلّل عدد الصفحات.")
        return txt


class ClaudeProvider:
    name = "Claude"

    def __init__(self, api_key: str, model: str, workspace_id: str = ""):
        self.api_key = api_key
        self.model = model
        self.workspace_id = (workspace_id or "").strip()
        self.on_debug = None

    def _client(self):
        try:
            import anthropic
        except ImportError:
            raise ProviderError("المكتبة غير مثبّتة — شغّل: pip install anthropic")
        kw = {"api_key": self.api_key}
        if self.workspace_id:
            kw["default_headers"] = {"anthropic-workspace-id": self.workspace_id}
        return anthropic.Anthropic(**kw)

    def check(self) -> str:
        try:
            self._client().messages.create(
                model=self.model, max_tokens=4,
                messages=[{"role": "user", "content": "hi"}])
        except Exception as e:
            if "workspace" in str(e).lower() and not self.workspace_id:
                raise ProviderError(
                    "المفتاح غير مرتبط بمساحة عمل.\n\n"
                    "الحل الأسرع: أنشئ مفتاحاً جديداً في console.anthropic.com\n"
                    "واختر مساحة عمل عند الإنشاء بدل Default.\n\n"
                    "أو الصق معرّف مساحة العمل في خانة «مساحة العمل» أدناه.")
            raise
        return "الاتصال ناجح"

    def list_models(self) -> list[str]:
        try:
            out = [m.id for m in self._client().models.list(limit=40)]
            return sort_models(out) if out else list(CLAUDE_MODELS)
        except Exception:
            return list(CLAUDE_MODELS)

    TOOL_NAME = "enregistrer_dossier_taxi"

    def extract(self, images: list[bytes], prompt: str):
        """يعيد dict مباشرة عبر الإخراج المُهيكل (tool use)."""
        client = self._client()
        content = [{
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg",
                       "data": base64.standard_b64encode(img).decode()},
        } for img in images]
        content.append({"type": "text", "text": prompt})

        tool = {
            "name": self.TOOL_NAME,
            "description": "تسجيل بيانات ملف رخصة استغلال سيارة أجرة المستخرجة من الوثائق.",
            "input_schema": build_response_schema("anthropic"),
        }
        msg = client.messages.create(
            model=self.model,
            max_tokens=16384,
            tools=[tool],
            tool_choice={"type": "tool", "name": self.TOOL_NAME},
            messages=[{"role": "user", "content": content}],
        )

        kinds = [getattr(b, "type", "?") for b in msg.content]
        stop = getattr(msg, "stop_reason", "?")
        if self.on_debug:
            self.on_debug(f"رد Claude: كتل={kinds or 'لا شيء'} · سبب التوقّف={stop}")

        for block in msg.content:
            if getattr(block, "type", "") == "tool_use":
                return dict(block.input)

        # احتياط: نموذج لم يستعمل الأداة
        texts = [b.text for b in msg.content if getattr(b, "type", "") == "text"]
        raw = "\n".join(texts).strip()
        if not raw:
            raise ProviderError(
                f"أعاد النموذج رداً فارغاً بلا بيانات.\n"
                f"الكتل المُستلمة: {kinds or 'لا شيء'} · سبب التوقّف: {stop}\n"
                f"إصدار مكتبة anthropic: {_pkg_version('anthropic')}\n\n"
                "إن كان الإصدار أقدم من 0.40 فحدّثه:  pip install -U anthropic")
        return raw




# ══════════════════════════════════════════════════════════════════════
#  المزوّد المحلي — Tesseract OCR بدون إنترنت
# ══════════════════════════════════════════════════════════════════════

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


# ══════════════════════════════════════════════════════════════════════
#  استخراج الحقول من نص OCR — يعمل صفحةً صفحة حسب نوع الوثيقة
#  (رخصة الاستغلال · البطاقة الرمادية · بطاقة التعريف · CNAS · CASNOS ·
#   شهادة التكوين · قرار المجاهدين · رخصة السياقة)
#  الصفحات مفصولة بالمحرف \f.  كل قيمة لا تُقرأ بثقة تُترك فارغة.
# ══════════════════════════════════════════════════════════════════════
import re as _re
import datetime as _dt
import difflib as _difflib

_BIDI = dict.fromkeys(map(ord, "‎‏‪‫‬‭‮"
                                "⁦⁧⁨⁩﻿​"), None)
_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

_COMMUNES_32 = [
    "البيض", "الأبيض سيد الشيخ", "بوعلام", "الرقاصة", "بريزينة", "الغاسول",
    "استيتن", "الشقيق", "بوقطب", "الكاف الأحمر", "عين العراك", "اربوات",
    "الشلالة", "بوسمغون", "المهرة", "الخيثر", "توسمولين", "سيدي سليمان",
    "الكراكدة", "سيدي عمر", "سيدي الطيفور", "سيدي عامر", "الأبيض سيدي الشيخ",
    "الكاف لحمر", "عين العراك",
]

# «تاريخ الميلاد» كما يشوّهه OCR أحياناً
_DOB = r"(?:الميلاد|المبلاد|اليلاد|المهلاد|الميلد|الملاد|شيلاد)"

_GROUPAGE = {"A+": "A +", "A-": "A -", "B+": "B + ", "B-": "B -",
             "AB+": "AB + ", "AB-": "AB -", "O+": "O +", "O-": "O -"}

_MARQUES = ["RENAULT", "DACIA", "HYUNDAI", "PEUGEOT", "CHEVROLET", "TOYOTA",
            "KIA", "SUZUKI", "VOLKSWAGEN", "CITROEN", "FIAT", "SEAT", "OPEL",
            "FORD", "NISSAN", "SKODA", "MAZDA", "MITSUBISHI", "GEELY", "CHERY",
            "DAEWOO", "HONDA", "ISUZU", "MERCEDES", "BMW", "AUDI", "LADA",
            "SSANGYONG", "JAC", "BYD", "DFSK", "BAIC", "DONGFENG"]
_MARQUES_AR = {"رونو": "RENAULT", "داسيا": "DACIA", "هيونداي": "HYUNDAI",
               "بيجو": "PEUGEOT", "شفروليه": "CHEVROLET", "شيفروليه": "CHEVROLET",
               "تويوتا": "TOYOTA", "كيا": "KIA", "سوزوكي": "SUZUKI"}


def _clean(t: str) -> str:
    return (t or "").translate(_BIDI).translate(_DIGITS)


def _ar(s: str) -> str:
    """تطبيع عربي للمقارنة فقط."""
    s = _re.sub(r"[ًٌٍَُِّْـ]", "", s or "")
    s = _re.sub(r"[أإآٱ]", "ا", s)
    s = s.replace("ة", "ه").replace("ى", "ي").replace("ؤ", "و").replace("ئ", "ي")
    return _re.sub(r"[^ء-ي]", "", s)


def _similar(a: str, b: str) -> float:
    a, b = _ar(a), _ar(b)
    if not a or not b:
        return 0.0
    return _difflib.SequenceMatcher(None, a, b).ratio()


def _is_ar_word(w: str) -> bool:
    return bool(w) and bool(_re.fullmatch(r"[ء-ي]{2,}", w))


def _ar_words(s: str, maxn: int = 3) -> str | None:
    """أول كلمات عربية متتالية (حتى maxn) — يتوقف عند أول رمز غريب."""
    out = []
    for w in _re.split(r"[\s:.,؛،]+", (s or "").strip()):
        if not w:
            continue
        if _is_ar_word(w):
            out.append(w)
            if len(out) >= maxn:
                break
        else:
            break
    return " ".join(out) or None


def _date(y, m, d, lo=1900, hi=2040):
    try:
        y, m, d = int(y), int(m), int(d)
        if lo <= y <= hi:
            return _dt.date(y, m, d).isoformat()
    except (ValueError, TypeError):
        pass
    return None


def _dates(s: str, lo=1900, hi=2040) -> list[str]:
    """كل التواريخ الصالحة في النص بالترتيب (YYYY.MM.DD أو DD.MM.YYYY)."""
    out = []
    for m in _re.finditer(r"(?<!\d)(\d{4})\s?[./\-]\s?(\d{1,2})\s?[./\-]\s?(\d{1,2})(?!\d)"
                          r"|(?<!\d)(\d{1,2})\s?[./\-]\s?(\d{1,2})\s?[./\-]\s?(\d{4})(?!\d)", s):
        if m.group(1):
            v = _date(m.group(1), m.group(2), m.group(3), lo, hi)
        else:
            v = _date(m.group(6), m.group(5), m.group(4), lo, hi)
        if v:
            out.append(v)
    return out


def _title(s: str | None) -> str | None:
    if not s:
        return None
    s = _re.sub(r"[^A-Za-z\-' ]", " ", s)
    s = " ".join(w for w in s.split() if len(w) >= 2)
    if not s or len(s) < 2:
        return None
    return " ".join(p.capitalize() for p in s.split())


def _latin_after(label_rx: str, page: str, maxw: int = 3) -> str | None:
    m = _re.search(label_rx + r"\s*[:.]?\s*((?:[A-Z][A-Za-z\-']+\s*){1,%d})" % maxw, page)
    if not m:
        return None
    stop = {"FILS", "FILLE", "ET", "DE", "NOM", "PRENOM", "DATE", "LIEU", "N", "NE", "NEE"}
    words = []
    for w in m.group(1).split():
        if w.upper() in stop or not _re.fullmatch(r"[A-Z][A-Z\-']+|[A-Z][a-z\-']+", w):
            break
        # بعد كلمة بحروف كبيرة لا نقبل إلا كلمات بحروف كبيرة (≥3) — الباقي ضجيج
        if words and words[0].isupper() and not (w.isupper() and len(w) >= 3):
            break
        words.append(w)
    return _title(" ".join(words[:maxw])) if words else None


def _nins(s: str) -> list[str]:
    out = []
    for line in s.split("\n"):
        j = _re.sub(r"[\s.]+", "", line)
        for m in _re.finditer(r"(?<!\d)([12][01]\d{16})(?!\d)", j):
            if m.group(1) not in out:
                out.append(m.group(1))
    return out


def _nins_loose(s: str) -> list[str]:
    """نوافذ 18 خانة داخل سلاسل من 19-20 رقماً (رقم زائد ألصقه OCR).
    لا تُستعمل إلا إذا وافقت سنة الميلاد المعروفة."""
    out = []
    for line in s.split("\n"):
        j = _re.sub(r"[\s.]+", "", line)
        for m in _re.finditer(r"(?<!\d)(\d{19,20})(?!\d)", j):
            run = m.group(1)
            for i in range(len(run) - 17):
                c = run[i:i + 18]
                if _re.match(r"[12][01]", c) and c not in out:
                    out.append(c)
    return out


def _nin_year(nin: str) -> int | None:
    try:
        return 1000 + int(nin[2:5])
    except (ValueError, TypeError):
        return None


def _commune(s: str) -> str | None:
    s = _ar_words(s, 4)
    if not s:
        return None
    best, score = None, 0.0
    for c in _COMMUNES_32:
        for k in range(1, 4):
            cand = " ".join(s.split()[:k])
            r = _similar(cand, c)
            if r > score:
                best, score = c, r
    return best if score >= 0.8 else None


def _parents(s: str, nom: str = ""):
    """«سليمان بن بوبكر و شريطات عائشة بنت ...» ← (الأب، لقب الأم، اسم الأم)."""
    s = _re.sub(r"[^ء-ي\s]", " ", s or "")
    parts = _re.split(r"\s+و\s*", s, maxsplit=1)
    pere = None
    if parts and parts[0].strip():
        w = parts[0].split()
        if len(w) > 1 and nom and _similar(w[0], nom) >= 0.8:
            w = w[1:]          # «راجع أحمد» ← اللقب ثم اسم الأب
        if w and w[0] in ("عبد", "ابو", "أبو") and len(w) > 1:
            pere = w[0] + " " + w[1]
        elif w and _is_ar_word(w[0]) and w[0] not in ("بن", "ابن"):
            pere = w[0]
    nom_m = pre_m = None
    if len(parts) > 1:
        w = parts[1].split()
        w = w[:w.index("بنت")] if "بنت" in w else w
        if w and w[0] in ("بن", "ابن") and len(w) > 2:
            w = [w[0] + " " + w[1]] + w[2:]
        if len(w) >= 2 and all(_is_ar_word(x.replace(" ", "")) for x in w[:2]):
            nom_m, pre_m = w[0], w[1]
    return pere, nom_m, pre_m


# ── تصنيف الصفحات ────────────────────────────────────────────────────
def _kind(p: str) -> set[str]:
    k = set()
    if _re.search(r"رخصة\s*مؤقتة|رقم\s*الباب", p):
        k.add("lic")
    if _re.search(r"التسلسلي|الحمولة|CARROSSERIE|GENRE\s+MARQUE|IMMATRICULATION|"
                  r"QUITTANCE|الرقم\s*السا|سنة\s*[اأ]ول\s*است", p):
        k.add("cg")
    if _re.search(r"بطا\S*\s*التعريف|سلطة\s*الإصدار|لطة\s*الإصدار", p):
        k.add("id")
    if _re.search(r"لغير\s*الأجراء|لغير\s*الاجراء|CASNOS|NON\s+SALARI", p):
        k.add("casnos")
    elif _re.search(r"للتأمينات|للتامينات|CNAS|الأجراء", p):
        k.add("cnas")
    if _re.search(r"دفتر\s*\S*\s*مقاعد|دقتر\s*مقاعد|الدورة\s*رقم", p):
        k.add("form")
    if _re.search(r"بصفت", p) and _re.search(r"مجاهد|شهيد|مكان\s*الاست", p):
        k.add("dec")
    if _re.search(r"رخصة\s*ال\S*ياقة|PERMIS\s+DE\s+CONDUIRE|DRIVING", p) and "lic" not in k:
        k.add("permis")
    return k


def _brand(s: str, short_ok: bool) -> str | None:
    """الصنف من نص (بحروف كبيرة). الأسماء القصيرة (KIA…) لا تُقبل إلا في
    سطر الصنف نفسه (short_ok) لأنها تظهر صدفةً في ضجيج OCR."""
    for mk in _MARQUES:
        if (short_ok or len(mk) >= 5) and _re.search(r"\b" + mk + r"\b", s):
            return mk
    # مطابقة متسامحة مع أخطاء OCR («RENANLT»، «CHEVR OLETE»)
    ws = _re.findall(r"\b[A-Z]{3,10}\b", s)
    best, score = None, 0.0
    for i, w in enumerate(ws):
        for cand in {w, w + ws[i + 1] if i + 1 < len(ws) else w}:
            if len(cand) < 5:
                continue
            for mk in _MARQUES:
                if len(mk) < 5:
                    continue
                r = _difflib.SequenceMatcher(None, cand, mk).ratio()
                if r > score:
                    best, score = mk, r
    return best if score >= 0.8 else None


def _carrosserie(s: str) -> str | None:
    """الهيكل كما هو مطبوع في البطاقة (BREAK / BERLINE / C,I) — توحيده مع
    قيم ملف Excel يتولاه فحص البرنامج بعد الاستخراج."""
    s = (s or "").upper()
    if _re.search(r"\bB?R[EI]?AK\b|BREAK", s):
        return "BREAK"
    if _re.search(r"BERLI|برلين", s):
        return "BERLINE"
    if _re.search(r"\bC\.?\s?I\b|ق\.?\s?د", s):
        return "C,I"
    return None


def _tsv_words(tsv: str) -> list[dict]:
    out = []
    for line in (tsv or "").split("\n"):
        c = line.split("\t")
        if len(c) < 12 or c[0] != "5":
            continue
        t = _clean(c[11]).strip()
        if not t:
            continue
        try:
            l, tp, w, h = int(c[6]), int(c[7]), int(c[8]), int(c[9])
        except ValueError:
            continue
        out.append({"l": l, "t": tp, "r": l + w, "b": tp + h, "h": h, "text": t,
                    "U": _re.sub(r"[^A-Z0-9]", "", t.upper()
                                 .replace("É", "E").replace("È", "E"))})
    return out


_CG_LABELS = {          # كلمة العنوان المطبوعة تحت القيمة ← الحقل
    "MARQUE": "marque", "PUISSANCE": "puissance", "ENERGIE": "energie",
    "PLACES": "n_siege", "ASSISES": "n_siege", "CARROSSERIE": "carrossere",
    "SERIE": "n_chasis", "ANNEE": "annee_premier_utilisation",
    "PREMIERE": "annee_premier_utilisation", "UTILE": "charge utile",
    "POIDS": "poids total", "TYPE": "type",
}
_LABEL_WORDS = set(_CG_LABELS) | {"GENRE", "DANS", "LA", "DU", "N", "TOTAL", "EN",
                                  "CHARGE", "DE", "MISE", "CIRCULATION", "NUMERO",
                                  "PRECEDENT", "DIMMATRICULATION", "IMMATRICULATION"}


def _cg_cells(tsv: str) -> dict:
    """قراءة جدول البطاقة الرمادية بالمواقع: لكل عنوان مطبوع (MARQUE، PUISSANCE…)
    نأخذ الكلمات التي فوقه مباشرة في عموده. كل كلمة تُنسب إلى أقرب عنوان في
    الصف نفسه، فلا تختلط قيمة خانة بجارتها."""
    ws = _tsv_words(tsv)
    if not ws:
        return {}

    # ① مراسي العناوين: (الحقل، الحدّ الأيسر، الأيمن، الأعلى، الارتفاع)
    anchors = []
    used = set()
    for lab in ws:
        U = lab["U"]
        if id(lab) in used:
            continue
        line = [w for w in ws if abs(w["t"] - lab["t"]) <= 15]
        f = None
        group = [lab]
        if U == "SERIE":
            f = "n_chasis"
            group = [w for w in line if (w["U"] in ("N", "DANS", "LA", "SERIE", "DU")
                                         and abs(w["l"] - lab["l"]) < 260)
                     or (w["U"] == "TYPE" and 0 < w["l"] - lab["l"] < 200)]
        elif U == "TYPE":
            ser = [w for w in line if w["U"] == "SERIE"]
            if ser and 0 < lab["l"] - ser[0]["l"] < 200:
                continue                         # جزء من «SERIE DU TYPE»
            f = "type"
        elif U in ("ANNEE", "PREMIERE"):
            f = "annee_premier_utilisation"
            group = [w for w in line if w["U"] in ("ANNEE", "DE", "PREMIERE")
                     and abs(w["l"] - lab["l"]) < 220]
        elif U in ("PLACES", "ASSISES"):
            f = "n_siege"
        elif U in ("MARQUE", "PUISSANCE", "ENERGIE", "CARROSSERIE"):
            f = _CG_LABELS[U]
        elif U in ("GENRE", "POIDS", "UTILE", "PRECEDENT", "DIMMATRICULATION",
                   "IMMATRICULATION"):
            f = "_" + U                          # مرساة بلا حقل: تحدّ الجيران فقط
        if not f:
            continue
        for w in group:
            used.add(id(w))
        if any(an["f"] == f for an in anchors):
            continue
        anchors.append({"f": f, "l": min(w["l"] for w in group),
                        "r": max(w["r"] for w in group), "t": min(w["t"] for w in group),
                        "h": max(w["h"] for w in group)})
    for an in anchors:
        an["cx"] = (an["l"] + an["r"]) / 2

    labelish = _LABEL_WORDS | {"PLACE", "ASSISE"}

    # ارتفاع سطر العناوين المطبوعة (≈17 نقطة عند تكبير 2) — مقياس الصفحة
    hs = sorted(an["h"] for an in anchors) or [18]
    hh = max(12, hs[len(hs) // 2])

    def value_of(an: dict) -> str | None:
        row = [o for o in anchors if abs(o["t"] - an["t"]) <= 2.5 * hh]
        lo = an["t"] - 5.5 * hh
        cands = []
        for w in ws:
            cy = (w["t"] + w["b"]) / 2       # القيمة بخط عريض قد تلامس العنوان
            if w["U"] in labelish or not (lo <= cy <= an["t"] + 0.5 * hh):
                continue
            cx = (w["l"] + w["r"]) / 2
            if not (an["l"] - 45 <= cx <= an["r"] + 45):
                continue
            nearest = min(row, key=lambda o: 0 if o["l"] <= cx <= o["r"] else
                          min(abs(cx - o["l"]), abs(cx - o["r"])))
            if nearest is an:
                cands.append(w)
        if not cands:
            return None
        cyc = lambda c: (c["t"] + c["b"]) / 2
        near = max(cyc(c) for c in cands)      # الأقرب إلى العنوان
        line = sorted([c for c in cands if near - cyc(c) <= max(20, 1.2 * hh)],
                      key=lambda c: c["l"])
        return " ".join(c["text"] for c in line)

    out: dict = {}
    for an in anchors:
        f = an["f"]
        if f.startswith("_"):
            continue
        v = value_of(an)
        if not v:
            continue
        VU = v.upper()
        V = _re.sub(r"[^A-Z0-9]", "", VU)
        if f == "marque":
            mk = _brand(VU, short_ok=True)
            if mk:
                out[f] = mk
        elif f == "type":
            if 5 <= len(V) <= 12 and _re.search(r"\d", V) and _re.search(r"[A-Z]", V):
                out[f] = V
        elif f == "n_chasis":
            if len(V) == 17:
                out[f] = V.translate(str.maketrans("IOQ", "100"))   # لا I/O/Q في VIN
            elif len(V) == 8 and len(_re.sub(r"\D", "", V)) >= 5:
                out[f] = V
        elif f == "puissance":
            d = _re.sub(r"\D", "", VU)
            if _re.fullmatch(r"0?\d{1,2}", d or "x") and 3 <= int(d) <= 20:
                out[f] = f"{int(d)} CHEVEAUX"
        elif f == "n_siege":
            d = _re.sub(r"\D", "", VU)
            if _re.fullmatch(r"0?[2-9]", d or "x"):
                out[f] = str(int(d))
        elif f == "energie":
            if _re.search(r"\bGO\b|GASOIL|مازوت", VU):
                out[f] = "Diesel"
            elif _re.search(r"GPL|غاز|\bGP\b", VU):
                out[f] = "GPL"
            elif _re.search(r"\bES\b|\bEG\b|ESS|بنزين", VU):
                out[f] = "Essence"
        elif f == "carrossere":
            c = _carrosserie(VU)
            if c:
                out[f] = c
        elif f == "annee_premier_utilisation":
            m = _re.search(r"(?<!\d)((?:19[89]|20[0-2])\d)(?!\d)", v)
            if m and int(m.group(1)) <= _dt.date.today().year:
                out[f] = m.group(1)
    # الطراز بداية الرقم التسلسلي عادةً: نصحّح به I/O
    fix = str.maketrans("IOQ", "100")
    if out.get("type") and len(out.get("n_chasis", "")) == 17 and \
            out["n_chasis"].startswith(out["type"].translate(fix)):
        out["type"] = out["type"].translate(fix)
    return out


def extract_taxi(text: str, wilaya: int = 32, tsv: str | None = None) -> dict:
    """text: نص الصفحات مفصولة بـ \\f.  tsv (اختياري): مخرجات Tesseract TSV
    للصفحات نفسها مفصولة بـ \\f — تتيح قراءة جدول البطاقة الرمادية بالمواقع."""
    pages = [_clean(p) for p in _re.split(r"\f", text or "") if p.strip()]
    if not pages:
        pages = [""]
    kinds = [_kind(p) for p in pages]

    def pg(kind):
        return [p for p, k in zip(pages, kinds) if kind in k]

    def first(rx, srcs, flags=0, grp=1):
        for s in srcs:
            m = _re.search(rx, s, flags)
            if m:
                return m.group(grp).strip()
        return None

    op, ch, ve, lm = {}, {}, {}, {}

    # ① رخصة الاستغلال المؤقتة ───────────────────────────────────────
    lic = pg("lic") or pages
    for p in lic:
        m = _re.search(r"اللقب\s*[:\-]\s*([^\n]+)", p)
        if m and not op.get("nom_ar"):
            op["nom_ar"] = _ar_words(m.group(1), 3)
        m = _re.search(r"(?:^|\n)[^\n]{0,4}ال[اإ]سم\s*[:\-]\s*([^\n]+)", p)
        if m and not op.get("prenom_ar"):
            op["prenom_ar"] = _ar_words(m.group(1), 3)
        m = _re.search(r"العنوان\s*[:\-]\s*([^\n]+)", p)
        if m and not op.get("adresse_ar"):
            op["adresse_ar"] = _ar_words(m.group(1), 5)
        m = _re.search(r"الصفة\s*[:\-]\s*(\S+)", p)
        if m and not op.get("qualite_oper_physique"):
            q = m.group(1)
            if "مكتر" in q or "مستأجر" in q:
                op["qualite_oper_physique"] = "Locatiare"
            elif _re.search(r"مستفيد|صاحب|مالك|ذو", q):
                op["qualite_oper_physique"] = "Bénéféciaire"
        m = _re.search(r"محيط[^\n:]*:\s*([^\n]+)", p) or \
            _re.search(r"سيارة\s*[اأ]جرة\s*(فردية|ما\s*بين\s*\S+)", p)
        if m and not op.get("nature_activite"):
            v = m.group(1)
            if "فردي" in v:
                op["nature_activite"] = "Individuelle"
            elif "الولاي" in v:
                op["nature_activite"] = "Inter Wilaya"
            elif "البلدي" in v:
                op["nature_activite"] = "Inter Communalle"
        # سطراً سطراً — بعض الأسطر تأتي معكوسة («2245 : رقم الباب»)
        for line in p.split("\n"):
            if _re.search(r"رقم\s*الباب", line) and not op.get("n-porte"):
                m = _re.search(r"رقم\s*الباب\s*[:\-]?\s*(\d{1,5})\b", line) or \
                    _re.search(r"(?<![\d/.])(\d{1,5})\s*[:\-]\s*رقم\s*الباب", line)
                if m:
                    op["n-porte"] = m.group(1)
            if _re.search(r"رقم\s*التسجيل", line) and not ve.get("n_immatriculation"):
                m = _re.search(r"(?<!\d)(\d{4,6})[.\-\s](\d{3})[.\-\s](\d{2})(?!\d)", line)
                if m:
                    ve["n_immatriculation"] = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            if _re.search(r"انتهاء|غاية", line) and not op.get("date_fin_contrat"):
                d = _dates(line, 2020, 2040)
                if d:
                    op["date_fin_contrat"] = d[0]
        m = _re.search(r"حرر\S*\s*ب\S*\s*في\s*[:.]*\s*([\d/.\-]{8,10})", p)
        if m and not op.get("date_debut_contrat"):
            d = _dates(m.group(1), 2020, 2040)
            if d:
                op["date_debut_contrat"] = d[0]
        # رقم الرخصة «الرقم: 43/2026» ← 2026/43
        m = _re.search(r"الر\s?قم[^\n\d]{0,8}([\d/ \t.|\]\[]{3,20})", p)
        if m and not ch.get("n_delivrance_licence_exploitation"):
            grp = _re.findall(r"\d+", m.group(1))
            year = next((g for g in grp if _re.fullmatch(r"202\d", g)), None)
            if year is None and grp and _re.fullmatch(r"202\d\d{1,4}", grp[0]):
                year, grp = grp[0][:4], [grp[0][:4], grp[0][4:]]
            rest = [g for g in grp if g != year]
            if year and len(rest) == 1 and 2 <= len(rest[0]) <= 4:
                ch["n_delivrance_licence_exploitation"] = f"{year}/{int(rest[0])}"

    # إن غابت رخصة الاستغلال: الاسم من شهادة CNAS / CASNOS
    for p in pg("cnas") + pg("casnos") + pg("id"):
        if not op.get("nom_ar"):
            op["nom_ar"] = _ar_words(first(r"الل?قب\s*[:\-]\s*([^\n]+)", [p]) or "", 3)
        if not op.get("prenom_ar"):
            op["prenom_ar"] = _ar_words(first(r"ال[اإ]سم\s*[:\-]\s*([^\n]+)", [p]) or "", 3)

    nom_op = op.get("nom_ar") or ""
    pre_op = op.get("prenom_ar") or ""

    def is_op_page(p: str) -> bool:
        """هل الصفحة تخصّ المتعامل؟ (لقبه أو اسمه الكامل مذكور فيها)"""
        if not nom_op:
            return True
        words = _re.findall(r"[ء-ي]{2,}", p)
        return any(_similar(w, nom_op) >= 0.75 for w in words) or \
            any(_similar(w, pre_op) >= 0.85 for w in words if len(_ar(w)) >= 4)

    # ② تاريخ الميلاد والأبوان والأسماء اللاتينية — من وثائق المتعامل
    dn_votes: dict[str, int] = {}
    op_docs = [p for p in pg("cg") + pg("cnas") + pg("casnos") + pg("form")]
    for p in op_docs:
        for m in _re.finditer(r"(?:" + _DOB + r"|المولود\S*\s*(?:ب\S*|في)|Naissance)[^\n]{0,60}", p):
            for d in _dates(m.group(0), 1940, 2007):
                dn_votes[d] = dn_votes.get(d, 0) + 2
    for p in pg("id"):
        if is_op_page(p):
            for m in _re.finditer(_DOB + r"[^\n]{0,30}", p):
                for d in _dates(m.group(0), 1940, 2007):
                    dn_votes[d] = dn_votes.get(d, 0) + 1

    # رقم التعريف الوطني: نفضّل ما يوافق سنة الميلاد
    all_nins = []
    for p in pages:
        for n in _nins(p):
            if n not in all_nins:
                all_nins.append(n)
    dn_op = max(dn_votes, key=lambda d: dn_votes[d]) if dn_votes else None
    nin_op = None
    if dn_op:
        nin_op = next((n for n in all_nins if _nin_year(n) == int(dn_op[:4])), None)
        if nin_op is None:
            loose = [n for p in pages for n in _nins_loose(p)
                     if _nin_year(n) == int(dn_op[:4])]
            if len(set(loose)) == 1:
                nin_op = loose[0]
    else:
        # بلا تاريخ ميلاد: نقبل الرقم فقط من بطاقة تعريف تحمل لقب المتعامل
        own = [n for p in pg("id") if is_op_page(p) for n in _nins(p)]
        if len(set(own)) == 1 and 1940 <= (_nin_year(own[0]) or 0) <= 2007:
            nin_op = own[0]
    if nin_op:
        op["nin"] = nin_op
    if dn_op:
        op["date_naissance"] = dn_op

    for p in pg("id"):
        if is_op_page(p) or (nin_op and nin_op in _re.sub(r"[\s.]", "", p)):
            m = _re.search(r"الجنس\s*[:.]?\s*(\S+)", p)
            if m and not op.get("sexe"):
                if _re.search(r"ذكر|زكر|ذكز", m.group(1)):
                    op["sexe"] = "M"
                elif _re.search(r"[اأ]نث", m.group(1)):
                    op["sexe"] = "F"
            m = _re.search(r"Rh\s*[:.]?\s*(AB|A|B|O|0)\s*([+\-#])", p)
            if m and not ch.get("groupage"):
                g = m.group(1).replace("0", "O") + ("-" if m.group(2) == "-" else "+")
                ch["groupage"] = _GROUPAGE.get(g)

    for p in pg("cnas") + pg("casnos") + pg("cg"):
        if not is_op_page(p):
            continue
        if not op.get("nom"):
            op["nom"] = _latin_after(r"Nom\s*", p, 2)
        if not op.get("prenom"):
            op["prenom"] = _latin_after(r"Pr[ée]nom\S*", p, 2)
        if not op.get("prenom_pere"):
            op["prenom_pere"] = _latin_after(r"Fils\s*(?:\(Fille\)\s*)?de", p, 2)
        m = _re.search(r"(?:ابن|أبن|اين|إن)\s*(?:\(ة\))?\s*:\s*([^\n]+)", p)
        m2 = _re.search(r"Et\s*de\s*[:.]?\s*([^\n]+)", p)
        if m and not op.get("prenom_pere_ar"):
            line = m.group(1)
            if m2 and " و " not in line:
                line = line + " و " + m2.group(1)
            pere, nm, pm = _parents(line, nom_op)
            if pere:
                op["prenom_pere_ar"] = pere
            if nm and not op.get("nom_mere_ar"):
                op["nom_mere_ar"], op["prenom_mere_ar"] = nm, pm

    # CG بالفرنسية «M. BENKHIRA Ali»
    if not op.get("nom"):
        for p in pg("cg"):
            m = _re.search(r"\bM[.\s]+([A-Z]{3,})\s+([A-Z][A-Za-z]{2,})", p)
            if m:
                op["nom"], op["prenom"] = _title(m.group(1)), _title(m.group(2))
                break

    # CNAS: رقم الضمان (12 خانة) ورقم عقد الميلاد
    for p in pg("cnas"):
        if not is_op_page(p):
            continue
        m = _re.search(r"(?:تحت\s*رقم|Sous\s*le\s*num\S*)\s*:?\s*([\d/\s]{10,16})", p)
        if m and not op.get("numero_cnas"):
            d = _re.sub(r"\D", "", m.group(1))
            if len(d) == 12 and not _re.search(r"غير\s*منتسب\s*(?:إلى|الى|ل)\s*هي", p[:m.start()]):
                op["numero_cnas"] = d
        m = _re.search(r"(?:عقد\s*الميلاد|N°\s*Acte)\s*:?\s*(\d{2,5})\b", p)
        if m and not op.get("n_acte_naissance"):
            op["n_acte_naissance"] = str(int(m.group(1)))
    for p in pg("casnos"):
        m = _re.search(r"\b(2323\d{7})\b", p)
        if m and not _re.search(r"غير\s*منتسب|NON\s*AFFILI|N.EST\s*PAS", p):
            op["numero_casnos"] = m.group(1)
    for p in pg("cnas") + pg("casnos"):
        m = _re.search(r"Fait\s*le\s*:?\s*(\d{2}/\d{2}/\d{4})", p)
        if m:
            d = _dates(m.group(1), 2023, 2030)
            if d and (not ch.get("date_mise_a_jour_securite_sociale")
                      or d[0] > ch["date_mise_a_jour_securite_sociale"]):
                ch["date_mise_a_jour_securite_sociale"] = d[0]

    # ③ البطاقة الرمادية ─────────────────────────────────────────────
    # أولاً: القراءة بالمواقع — القيمة هي ما فوق العنوان المطبوع في الخانة نفسها
    tsv_pages = tsv.split("\f") if tsv else []
    for i, k in enumerate(kinds):
        if "cg" in k and i < len(tsv_pages) and tsv_pages[i].strip():
            for f, v in _cg_cells(tsv_pages[i]).items():
                ve.setdefault(f, v)

    for p in pg("cg"):
        up = p.upper()
        # الصنف: نبحث في سطر القيم فوق «GENRE MARQUE» فقط — كلمة «KIA» وسط
        # ضجيج الصفحة لا تعني أن المركبة كيا
        ls = up.split("\n")
        idx = [i for i, l in enumerate(ls) if "MARQUE" in l or "الصنف" in l]
        zone = "\n".join("\n".join(ls[max(0, i - 2): i + 1]) for i in idx)
        if not ve.get("marque") and zone:
            ve["marque"] = _brand(zone, short_ok=True)
        if not ve.get("marque"):
            ve["marque"] = _brand(up, short_ok=False)
        if not ve.get("marque"):
            ve.pop("marque", None)
        if not ve.get("marque"):
            for ar, mk in _MARQUES_AR.items():
                if ar in p:
                    ve["marque"] = mk
                    break
        # الطراز + الرقم التسلسلي (VIN) على السطر نفسه بعد الصنف
        if ve.get("marque"):
            m = _re.search(ve["marque"][:4] + r"\S*\s*[|]?\s*([A-Z0-9]{5,12})\s*[|]?\s*([A-Z0-9]{8,17})?", up)
            if m:
                typ, vin = m.group(1), m.group(2)
                if vin and len(vin) == 17:
                    fix = str.maketrans("IOQ", "100")      # لا I ولا O ولا Q في VIN
                    vin = vin.translate(fix)
                    if typ and vin.startswith(typ.translate(fix)):
                        typ = typ.translate(fix)
                if typ and _re.search(r"[A-Z]", typ) and _re.search(r"\d", typ):
                    ve.setdefault("type", typ)
                if vin and len(vin) in (8, 17) and _re.search(r"\d", vin):
                    ve.setdefault("n_chasis", vin)
        # الهيكل · الطاقة · القوة · المقاعد
        m = _re.search(r"\b(BREAK|BERLINE|BRIAK|REAK|BERLIN|C\.?I)\b[,.]?\s*(EG|ES|GO|GPL|ESS\S*)?"
                       r"\s*(\d{1,2})?\s*(\d)?", up)
        c = _carrosserie(m.group(1)) if m else \
            ("BERLINE" if "برلين" in p else "C,I" if _re.search(r"ق\s?\.\s?د", p) else None)
        if c:
            ve.setdefault("carrossere", c)
        if m and m.group(2):
            e = m.group(2)
            ve.setdefault("energie", "Diesel" if e == "GO" else
                          "GPL" if e == "GPL" else "Essence")   # EG = Essence (كما في ملفاتك)
            if m.group(3) and 3 <= int(m.group(3)) <= 15:
                ve.setdefault("puissance", f"{int(m.group(3))} CHEVEAUX")
            if m.group(4) and 2 <= int(m.group(4)) <= 9:
                ve.setdefault("n_siege", m.group(4))
        if not ve.get("energie"):
            if _re.search(r"بنز\S*\s*\+\s*غاز|\+\s*غاز", p):
                ve["energie"] = "GPL"
            elif _re.search(r"مازوت|\bGO\b|GASOIL", p):
                ve["energie"] = "Diesel"
        if _re.search(r"سياح|س\s?[.,]?\s?خ\b|\bVP\b", p):
            ve.setdefault("genre", "س,خ")
        # سنة أول استعمال: سنة 19xx/20xx بجانب رقمي التسجيل
        for m in _re.finditer(r"(?<!\d)((?:19[89]|20[0-2])\d)(?!\d)", p):
            ctx = p[max(0, m.start() - 40): m.end() + 40]
            if _re.search(r"\d{4,6}-\d{2,3}-?\d{1,2}", ctx) and int(m.group(1)) <= _dt.date.today().year:
                ve.setdefault("annee_premier_utilisation", m.group(1))
                break
        if not ve.get("n_immatriculation"):
            for m in _re.finditer(r"(?<!\d)(\d{5,6})-(\d{3})-(32)(?!\d)", p):
                ve["n_immatriculation"] = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                break
        m = _re.search(r"(?:ابن|أبن)\s*:\s*([^\n]+)", p)
        if m and is_op_page(p) and not op.get("prenom_pere_ar"):
            pere, nm, pm = _parents(m.group(1), nom_op)
            if pere:
                op["prenom_pere_ar"] = pere
            if nm:
                op["nom_mere_ar"], op["prenom_mere_ar"] = nm, pm

    # ④ شهادة التكوين ───────────────────────────────────────────────
    for p in pg("form"):
        head = "\n".join(l for l in p.split("\n")[:12]
                         if not _re.search(r"مرسوم|المرصوم|المرسوع|قرار|مقتضى", l))
        m = _re.search(r"رق\S*\s*\D{0,3}(\d{2,5})\b[^\n]*?\b(202\d)\b", head)
        if m and not ch.get("n_certificat_formation"):
            ch["n_certificat_formation"] = f"{m.group(2)}/{int(m.group(1))}"
        for m in _re.finditer(r"(?:البي\S*|في)\s*[:;,.]?\s*\S{0,6}\s*[;:,]?\s*"
                              r"(\d{4}[/\-]\d{2}[/\-]\d{2})", head):
            d = _dates(m.group(1), 2015, 2035)
            if d:
                ch.setdefault("date_certificat_formation", d[0])
                break

    # ⑤ رخصة السياقة (نادراً ما تُقرأ — صورة بطاقة) ───────────────────
    for p in pg("permis"):
        m = _re.search(r"\b(L\s?\d{8})\b|\b(\d{2}/\d{2}/\d{5})\b", p)
        if m and not ch.get("n_permis"):
            ch["n_permis"] = (m.group(1) or m.group(2)).replace(" ", "")

    # ⑥ قرار منح الرخصة لذي الحق (المجاهد/الأرملة/الابنة) ──────────────
    _STOP = {"ابنة", "إبنة", "ابن", "إبن", "بنت", "ابت", "ابئة", "المولود", "المولودة",
             "بصفتها", "بصفته", "واليا", "والي", "أرملة", "ارملة", "مكان"}
    _name_rx = (r"للسيد(?:\(?ة\)?|زة\)|ر\(ة\)|ه)?\s*[:؛)]*\s*"
                r"([ء-ي]{2,})\s+([ء-ي]{2,})(?:\s+([ء-ي]{2,}))?")
    for p in pg("dec"):
        all_l = p.split("\n")
        lines = [(i, l) for i, l in enumerate(all_l)
                 if "للسيد" in l and not _re.search(r"تعيين|زعيين|واليا", l)]
        # الأولوية لسطر «المادة الأولى: تمنح رخصة ... للسيد(ة)» ثم متن القرار ثم العنوان
        lines.sort(key=lambda il: (0 if _re.search(r"ادة\s*الا?و?لى|تمنح|تمتح|تملع", il[1])
                                   else 1 if il[0] >= 8 else 2, il[0]))
        lines = [l for _, l in lines]
        for l in lines:
            m = _re.search(_name_rx, l)
            if not m:
                continue
            nom_b, w2, w3 = m.group(1), m.group(2), m.group(3)
            if nom_b in _STOP or w2 in _STOP:
                continue
            pre_b = w2 + " " + w3 if (w2 in ("عبد", "أبو", "ابو") and w3) else w2
            lm.setdefault("nom_ar", nom_b)
            lm.setdefault("prenom_ar", pre_b)
            break
        m = _re.search(r"المولود\S*\s*بتاريخ\s*[:]?\s*([\d/.\-]{8,10})", p)
        if m:
            d = _dates(m.group(1), 1900, 2007)
            if d:
                lm.setdefault("date_naissance", d[0])
        m = _re.search(r"مكان\s*الاست\S*\s*(?:ب)?\S*\s+([^\n؛;]+)", p)
        if m:
            c = _commune(m.group(1))
            if c:
                lm.setdefault("concentration_exploitation_licence_modjahid", c)
        for l in p.split("\n")[:12]:
            if _re.search(r"مقتضى|بناء|بتاء|على\s*القرار", l):
                continue           # إحالة إلى نص قانوني آخر — ليس رقم هذا القرار
            m = _re.search(r"قرار\s*رقم\s*:?\s*(\d{2,5})(?![\d/])(?!\s+\d)", l)
            if m:
                lm.setdefault("n_decision_octroi", m.group(1))
                d = _dates(l[m.end():], 1962, 2030)
                if d:
                    lm.setdefault("date_decision_octroi", d[0])
                break
        if _re.search(r"تمنح|منح", p):
            lm.setdefault("type_licence", "Délivrée")

    # بطاقة ذي الحق: الرقم الوطني الذي يوافق اسمه أو سنة ميلاده
    if lm:
        by = int(lm["date_naissance"][:4]) if lm.get("date_naissance") else None
        for p in pg("id"):
            if is_op_page(p) and not (lm.get("nom_ar") and lm["nom_ar"] in p):
                continue
            ns = [n for n in _nins(p) if n != nin_op]
            for n in ns:
                if by is None or _nin_year(n) == by:
                    lm.setdefault("nin_modjahid", n)
            if lm.get("nin_modjahid") and not lm.get("date_naissance"):
                for m in _re.finditer(_DOB + r"[^\n]{0,30}", p):
                    d = _dates(m.group(0), 1900, 2007)
                    if d and int(d[0][:4]) == _nin_year(lm["nin_modjahid"]):
                        lm["date_naissance"] = d[0]
                        break
        if not lm.get("nin_modjahid"):
            acte_op = op.get("n_acte_naissance")

            def plausible(n):
                if n == nin_op:
                    return False
                if acte_op and str(int(n[9:14])) == acte_op:
                    return False   # رقم المتعامل نفسه مقروءاً بخطأ
                if by is not None:
                    return _nin_year(n) == by
                return n[1] == "1" or (_nin_year(n) or 3000) <= 1975
            rest = [n for n in all_nins if plausible(n)]
            if len(rest) == 1:
                lm["nin_modjahid"] = rest[0]

    # ⑦ التجميع والنسخ بين الأقسام ───────────────────────────────────
    porte = op.get("n-porte")
    nature = op.get("nature_activite")
    op.update({"id_wilaya": str(wilaya), "activite": "Transports par Taxi"})
    if op.get("date_debut_contrat") and op.get("date_fin_contrat") and \
            op["date_debut_contrat"] >= op["date_fin_contrat"]:
        op.pop("date_debut_contrat")

    ch.update({
        "id_wilaya": str(wilaya), "activite": "Transports par Taxi",
        "nature_activite": nature, "qualite_chauffeur": "Principal",
        "nin": op.get("nin"), "nom_ar": op.get("nom_ar"), "prenom_ar": op.get("prenom_ar"),
        "nom": op.get("nom"), "prenom": op.get("prenom"),
        "date_naissance": op.get("date_naissance"), "sexe": op.get("sexe"),
        "n_acte_naissance": op.get("n_acte_naissance"),
        "numero_cnas": op.get("numero_cnas"), "numero_casnos": op.get("numero_casnos"),
        "adresse_ar": op.get("adresse_ar"),
    })
    if ch.get("n_delivrance_licence_exploitation") and op.get("date_debut_contrat"):
        ch.setdefault("date_delivrance_licence_exploitation", op["date_debut_contrat"])

    ve.update({
        "id_wilaya": str(wilaya), "activite": "Transports par Taxi",
        "nature_activite": nature, "n_porte": porte,
        "adresse_ar": op.get("adresse_ar"),
    })
    if ch.get("n_delivrance_licence_exploitation"):
        ve.setdefault("n_certificat_cerculation", ch["n_delivrance_licence_exploitation"])
    if op.get("date_fin_contrat"):
        ve.setdefault("date_fin_validite_certificat", op["date_fin_contrat"])

    if lm:
        lm.update({"id_wilaya": str(wilaya), "n_porte": porte})

    def tidy(d):
        return {k: v for k, v in d.items() if v not in (None, "", [])}

    return {
        "operateur": tidy(op),
        "chauffeur": tidy(ch),
        "vehicule": tidy(ve),
        "licence_modjahid": tidy(lm),
    }


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


# ══════════════════════════════════════════════════════════════════════
#  مزوّد OCR.space — قراءة سحابية عبر ocr_space.py (requests فقط)
# ══════════════════════════════════════════════════════════════════════

OCRSPACE_MODELS = ["engine-2", "engine-3"]


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
    Tesseract، لتعمل قراءة خانات البطاقة الرمادية بالمواقع مع المزوّدَين."""
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
    (language=auto)، ثم الاستخراج نفسه المستعمل مع Tesseract."""
    name = "OCR.space"
    on_debug = None

    def __init__(self, key: str = "", model: str = "", *_, **__):
        self.key = (key or "").strip()
        self.model = model or OCRSPACE_MODELS[0]

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
        return json.dumps(extract_taxi("\f".join(texts), WILAYA_ID, tsv="\f".join(tsvs)),
                          ensure_ascii=False)


PROVIDERS = {"Gemini": GeminiProvider, "Claude": ClaudeProvider,
             "Offline": OfflineProvider, "OCR.space": OcrSpaceProvider}


def _strip_comments(txt: str) -> str:
    """إزالة تعليقات // و /* */ خارج السلاسل النصية."""
    out = []
    i, n = 0, len(txt)
    in_str = False
    while i < n:
        c = txt[i]
        if in_str:
            out.append(c)
            if c == "\\" and i + 1 < n:
                out.append(txt[i + 1]); i += 2; continue
            if c == '"':
                in_str = False
            i += 1; continue
        if c == '"':
            in_str = True; out.append(c); i += 1; continue
        if c == "/" and i + 1 < n:
            if txt[i + 1] == "/":
                while i < n and txt[i] != "\n":
                    i += 1
                continue
            if txt[i + 1] == "*":
                i += 2
                while i + 1 < n and not (txt[i] == "*" and txt[i + 1] == "/"):
                    i += 1
                i += 2; continue
        out.append(c); i += 1
    return "".join(out)


def _strip_trailing_commas(txt: str) -> str:
    out = []
    i, n = 0, len(txt)
    in_str = False
    while i < n:
        c = txt[i]
        if in_str:
            out.append(c)
            if c == "\\" and i + 1 < n:
                out.append(txt[i + 1]); i += 2; continue
            if c == '"':
                in_str = False
            i += 1; continue
        if c == '"':
            in_str = True; out.append(c); i += 1; continue
        if c == ",":
            j = i + 1
            while j < n and txt[j] in " \t\r\n":
                j += 1
            if j < n and txt[j] in "}]":
                i += 1; continue
        out.append(c); i += 1
    return "".join(out)


def parse_json_response(raw: str) -> dict:
    """تحليل رد النموذج مع إصلاح الأعطاب الشائعة."""
    txt = (raw or "").strip()

    # إزالة أسوار الشيفرة
    if txt.startswith("```"):
        body = txt[3:]
        if body.lstrip().lower().startswith("json"):
            body = body.lstrip()[4:]
        txt = body.split("```")[0] if "```" in body else body

    # اقتصاص ما بين أول { وآخر }
    start, end = txt.find("{"), txt.rfind("}")
    if start != -1 and end > start:
        txt = txt[start:end + 1]

    for attempt in (
        lambda s: s,
        _strip_comments,
        lambda s: _strip_trailing_commas(_strip_comments(s)),
    ):
        try:
            data = json.loads(attempt(txt))
            if isinstance(data, dict):
                return data
        except json.JSONDecodeError:
            continue

    raise json.JSONDecodeError("تعذّر إصلاح رد النموذج", txt[:400], 0)


# ══════════════════════════════════════════════════════════════════════
#  التحقّق من صحّة البيانات
# ══════════════════════════════════════════════════════════════════════

# حقول رقمية بحتة — أي حرف فيها خطأ قراءة ضوئية
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


# ══════════════════════════════════════════════════════════════════════
#  الاشتقاق — ملء ما لم يُقرأ من معلومات متاحة أصلاً
#
#  بنية رقم التعريف الوطني الجزائري (18 رقماً)، مستخرجة بتحليل
#  3510 سجلاً من قاعدة الولاية:
#     الخانة   1      : سلسلة الإصدار
#     الخانة   2      : الجنس   — 0 ذكر · 1 أنثى      (دقّة 100%)
#     الخانات  3–5    : سنة الميلاد (آخر ثلاثة أرقام) (دقّة 90%)
#     الخانات  6–9    : مكتب الحالة المدنية            (80% أحادي الولاية)
#     الخانات 10–14   : رقم شهادة الميلاد               (دقّة 96%)
# ══════════════════════════════════════════════════════════════════════

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


    return warn


# ══════════════════════════════════════════════════════════════════════
#  الكتابة في Excel
# ══════════════════════════════════════════════════════════════════════

def excel_column_map(ws) -> dict[str, int]:
    row1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    out = {}
    for ci, cell in enumerate(row1, 1):
        if cell is not None:
            out[str(cell).strip()] = ci
    return out


def _norm(s: str) -> str:
    return str(s).replace("-", "_").replace(" ", "_").strip().lower()


def find_column(col_map: dict[str, int], field: str) -> int | None:
    if field in col_map:
        return col_map[field]
    target = _norm(field)
    for k, ci in col_map.items():
        if _norm(k) == target:
            return ci
    return None


def next_empty_row(ws, ncols: int = 0) -> int:
    """
    أول صف فارغ تماماً.
    ⚠️ لا يكفي فحص العمود الأول: سجلّات الشركات تترك id_wilaya فارغاً
    بينما تملأ عشرين عموداً أخرى — فحصُ عمود واحد يدمّرها.
    """
    ncols = ncols or ws.max_column
    for ri in range(3, ws.max_row + 2):
        if all(ws.cell(row=ri, column=ci).value in (None, "")
               for ci in range(1, ncols + 1)):
            return ri
    return ws.max_row + 1


def duplicate_row(ws, col_map, chains, data) -> tuple[int | None, str]:
    """
    يبحث عن صف مطابق باستعمال سلاسل مفاتيح مرتّبة.
    تُجرَّب كل سلسلة بالترتيب؛ تُتخطّى أي سلسلة ينقص أحد حقولها.
    يعيد (رقم الصف أو None، وصف المفتاح المستعمل).
    """
    if isinstance(chains, str):
        chains = [(chains,)]
    for chain in chains or []:
        cols, vals = [], []
        for f in chain:
            ci = find_column(col_map, f)
            v = data.get(f)
            if not ci or v in (None, "", "null"):
                cols = []
                break
            cols.append(ci)
            vals.append(str(v).strip())
        if not cols:
            continue
        for ri in range(3, ws.max_row + 1):
            if all(ws.cell(row=ri, column=c).value is not None
                   and str(ws.cell(row=ri, column=c).value).strip() == t
                   for c, t in zip(cols, vals)):
                return ri, "+".join(chain)
    return None, ""


def write_to_excel(excel_path: str, data: dict, overwrite_dups: bool, log) -> dict:
    """كتابة البيانات في Excel. يعيد ملخّص العملية."""
    import openpyxl

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = str(Path(excel_path).with_name(f"{Path(excel_path).stem}_backup_{ts}.xlsx"))
    shutil.copy2(excel_path, backup)
    log(f"💾 نسخة احتياطية: {Path(backup).name}")

    wb = openpyxl.load_workbook(excel_path)
    summary = {"added": [], "updated": [], "skipped": [], "missing": [], "backup": backup}

    for gkey, grp in SCHEMA.items():
        row_data = {k: v for k, v in (data.get(gkey) or {}).items()
                    if v not in (None, "", "null")}
        if not row_data:
            continue
        # قسم لا يحمل إلا ثوابت مشتقّة ليس سجلاً — لا يُكتب
        if not section_active(data, gkey):
            log(f"⏭️  [{grp['sheet']}] لا معلومة جوهرية — لم يُكتب")
            continue
        sheet = grp["sheet"]
        if sheet not in wb.sheetnames:
            summary["missing"].append(sheet)
            log(f"⚠️  الورقة «{sheet}» غير موجودة — تم التخطي")
            continue

        ws = wb[sheet]
        col_map = excel_column_map(ws)

        dup_ri, by = duplicate_row(ws, col_map, grp.get("dup") or grp.get("key"), row_data)

        if dup_ri and not overwrite_dups:
            summary["skipped"].append((sheet, dup_ri))
            log(f"⏭️  [{sheet}] سجل مطابق في الصف {dup_ri} (بـ {by}) — تم التخطي")
            continue
        if not dup_ri and not any(
                find_column(col_map, f) and row_data.get(f) not in (None, "", "null")
                for chain in (grp.get("dup") or []) for f in chain):
            summary.setdefault("nokey", []).append(sheet)
            log(f"⚠️  [{sheet}] لا مفتاح صالح لكشف التكرار — قد يتكرّر السجل")

        ri = dup_ri or next_empty_row(ws)
        written = 0
        for field, value in row_data.items():
            ci = find_column(col_map, field)
            if ci:
                ws.cell(row=ri, column=ci).value = value
                written += 1
        if dup_ri:
            summary["updated"].append((sheet, ri, written))
            log(f"♻️  [{sheet}] تحديث الصف {ri} — {written} حقل")
        else:
            summary["added"].append((sheet, ri, written))
            log(f"✅ [{sheet}] إضافة في الصف {ri} — {written} حقل")

    wb.save(excel_path)
    log(f"💾 تم حفظ: {Path(excel_path).name}")
    return summary


# ══════════════════════════════════════════════════════════════════════
#  الطابور
# ══════════════════════════════════════════════════════════════════════

ST_PENDING  = "بانتظار"
ST_WORKING  = "قيد المعالجة"
ST_REVIEW   = "بانتظار المراجعة"
ST_SAVED    = "مُدخَل"
ST_SKIPPED  = "مكرّر"
ST_FAILED   = "فشل"

ST_COLOR = {
    ST_PENDING: "#64748b", ST_WORKING: "#1d4ed8", ST_REVIEW: "#b45309",
    ST_SAVED:   "#047857", ST_SKIPPED: "#7c3aed", ST_FAILED: "#b91c1c",
}

SRC_LABEL = {"manual": "يدوي", "folder": "مجلّد", "email": "بريد"}


class Job:
    """ملفٌّ واحد في رحلته عبر خطّ المعالجة."""
    _seq = 0

    def __init__(self, path, source="manual", note=""):
        Job._seq += 1
        self.id = Job._seq
        self.path = Path(path)
        self.source = source
        self.note = note
        self.status = ST_PENDING
        self.data: dict | None = None
        self.warnings: list[str] = []
        self.error = ""
        self.summary = ""
        self.created = datetime.datetime.now()

    @property
    def name(self):
        return self.path.name

    def label(self):
        if self.data:
            o = self.data.get("operateur") or {}
            nm = f"{o.get('nom_ar') or ''} {o.get('prenom_ar') or ''}".strip()
            pt = o.get("n-porte") or ""
            return f"{nm}  ·  باب {pt}" if nm or pt else self.name
        return self.name


class JobQueue:
    """طابور آمن للخيوط مع إشعار عند كل تغيّر."""

    def __init__(self, on_change=None):
        self._lock = threading.Lock()
        self._jobs: list[Job] = []
        self._seen: set[str] = set()
        self.on_change = on_change

    def _touch(self):
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass

    def add(self, path, source="manual", note="") -> "Job | None":
        key = str(Path(path).resolve()).lower()
        with self._lock:
            if key in self._seen:
                return None
            self._seen.add(key)
            job = Job(path, source, note)
            self._jobs.append(job)
        self._touch()
        return job

    def next_pending(self) -> "Job | None":
        with self._lock:
            for j in self._jobs:
                if j.status == ST_PENDING:
                    j.status = ST_WORKING
                    return j
        return None

    def update(self, job, **kw):
        with self._lock:
            for k, v in kw.items():
                setattr(job, k, v)
        self._touch()

    def all(self) -> list:
        with self._lock:
            return list(self._jobs)

    def get(self, jid) -> "Job | None":
        with self._lock:
            return next((j for j in self._jobs if j.id == jid), None)

    def counts(self) -> dict:
        with self._lock:
            c = {}
            for j in self._jobs:
                c[j.status] = c.get(j.status, 0) + 1
            return c

    def clear_done(self):
        with self._lock:
            self._jobs = [j for j in self._jobs
                          if j.status in (ST_PENDING, ST_WORKING, ST_REVIEW)]
        self._touch()


# ══════════════════════════════════════════════════════════════════════
#  مراقبة المجلّد
# ══════════════════════════════════════════════════════════════════════

class FolderWatcher:
    """
    يراقب مجلّداً ويضيف كل PDF جديد إلى الطابور.
    يتجاهل الملفات التي ما زالت قيد النسخ (حجمها يتغيّر).
    """

    def __init__(self, folder, queue, log, interval=6):
        self.folder = Path(folder)
        self.queue = queue
        self.log = log
        self.interval = interval
        self._stop = threading.Event()
        self._thread = None
        self._sizes: dict[str, int] = {}

    @property
    def running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.running:
            return
        self.folder.mkdir(parents=True, exist_ok=True)
        (self.folder / "تم_إدخالها").mkdir(exist_ok=True)
        (self.folder / "تحتاج_مراجعة").mkdir(exist_ok=True)
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        self.log(f"👁 المراقبة تعمل على: {self.folder}", "ok")

    def stop(self):
        self._stop.set()
        self.log("⏸ أُوقفت مراقبة المجلّد.", "warn")

    def scan_once(self) -> int:
        """مسحة واحدة — تعيد عدد الملفات الجديدة."""
        n = 0
        try:
            for p in sorted(self.folder.glob("*.pdf")):
                if not p.is_file():
                    continue
                key = str(p.resolve()).lower()
                size = p.stat().st_size
                # ننتظر استقرار الحجم حتى لا نقرأ ملفاً قيد النسخ
                if self._sizes.get(key) != size:
                    self._sizes[key] = size
                    continue
                if self.queue.add(p, "folder"):
                    n += 1
                    self.log(f"📥 ملف جديد: {p.name}", "info")
        except Exception as e:
            self.log(f"⚠️ خطأ في المراقبة: {e}", "warn")
        return n

    def _loop(self):
        while not self._stop.is_set():
            self.scan_once()
            self._stop.wait(self.interval)


# ══════════════════════════════════════════════════════════════════════
#  البريد الإلكتروني (IMAP)
# ══════════════════════════════════════════════════════════════════════

IMAP_PRESETS = {
    "Gmail":          ("imap.gmail.com", 993),
    "Outlook/Hotmail": ("outlook.office365.com", 993),
    "Yahoo":          ("imap.mail.yahoo.com", 993),
    "آخر":            ("", 993),
}


class MailFetcher:
    """
    يتّصل بصندوق البريد عبر IMAP، ينزّل مرفقات PDF من الرسائل غير المقروءة،
    ويضعها في مجلّد المراقبة — فتدخل خطّ المعالجة نفسه.
    """

    def __init__(self, host, port, user, password, folder="INBOX",
                 subject_filter="", sender_filter="", mark_read=True):
        self.host, self.port = host, int(port or 993)
        self.user, self.password = user, password
        self.mailbox = folder or "INBOX"
        self.subject_filter = (subject_filter or "").strip()
        self.sender_filter = (sender_filter or "").strip()
        self.mark_read = mark_read

    def _connect(self):
        import imaplib
        if not self.host:
            raise ProviderError("خادم البريد غير محدّد.")
        if not self.user or not self.password:
            raise ProviderError("بريد المستخدم أو كلمة المرور غير مُدخلة.")
        try:
            m = imaplib.IMAP4_SSL(self.host, self.port, timeout=30)
        except Exception as e:
            raise ProviderError(f"تعذّر الاتصال بـ {self.host}:{self.port} — {e}")
        try:
            m.login(self.user, self.password)
        except Exception as e:
            msg = str(e)
            hint = ""
            if "gmail" in self.host.lower():
                hint = ("\n\nGmail يرفض كلمة المرور العادية. أنشئ «كلمة مرور التطبيقات» من:\n"
                        "myaccount.google.com ← الأمان ← التحقّق بخطوتين ← كلمات مرور التطبيقات")
            raise ProviderError(f"فشل تسجيل الدخول: {msg}{hint}")
        return m

    def check(self) -> str:
        m = self._connect()
        try:
            typ, data = m.select(self.mailbox, readonly=True)
            if typ != "OK":
                raise ProviderError(f"تعذّر فتح الصندوق «{self.mailbox}»")
            total = int(data[0]) if data and data[0] else 0
            typ, d = m.search(None, "UNSEEN")
            unseen = len(d[0].split()) if typ == "OK" and d and d[0] else 0
            return f"الاتصال ناجح — {total} رسالة، منها {unseen} غير مقروءة"
        finally:
            try:
                m.logout()
            except Exception:
                pass

    def _criteria(self):
        c = ["UNSEEN"]
        if self.sender_filter:
            c += ["FROM", f'"{self.sender_filter}"']
        if self.subject_filter:
            c += ["SUBJECT", f'"{self.subject_filter}"']
        return c

    def fetch_pdfs(self, dest_dir, log, limit=25) -> list:
        """ينزّل مرفقات PDF ويعيد قائمة (المسار، وصف الرسالة)."""
        import email
        from email.header import decode_header, make_header

        dest = Path(dest_dir)
        dest.mkdir(parents=True, exist_ok=True)
        out = []
        m = self._connect()
        try:
            if m.select(self.mailbox)[0] != "OK":
                raise ProviderError(f"تعذّر فتح الصندوق «{self.mailbox}»")
            typ, data = m.search(None, *self._criteria())
            ids = data[0].split() if typ == "OK" and data and data[0] else []
            if not ids:
                log("📭 لا رسائل جديدة مطابقة.", "dim")
                return out
            log(f"📬 {len(ids)} رسالة جديدة.", "info")

            for num in ids[-limit:]:
                typ, raw = m.fetch(num, "(RFC822)")
                if typ != "OK" or not raw or not raw[0]:
                    continue
                msg = email.message_from_bytes(raw[0][1])

                def hdr(h):
                    try:
                        return str(make_header(decode_header(msg.get(h, "") or "")))
                    except Exception:
                        return msg.get(h, "") or ""

                subj, frm = hdr("Subject"), hdr("From")
                found = 0
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    fname = part.get_filename()
                    if not fname:
                        continue
                    try:
                        fname = str(make_header(decode_header(fname)))
                    except Exception:
                        pass
                    if not fname.lower().endswith(".pdf"):
                        continue
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                    safe = "".join(ch for ch in fname
                                   if ch not in '\\/:*?"<>|').strip() or "mail.pdf"
                    target = dest / safe
                    i = 1
                    while target.exists():
                        target = dest / f"{Path(safe).stem}_{i}{Path(safe).suffix}"
                        i += 1
                    target.write_bytes(payload)
                    out.append((target, f"{frm} — {subj}"))
                    found += 1
                    log(f"   📎 {target.name}  ({frm})", "ok")

                if found and self.mark_read:
                    m.store(num, "+FLAGS", "\\Seen")
                elif not found:
                    log(f"   ✉️ بلا مرفق PDF: {subj[:50]}", "dim")
            return out
        finally:
            try:
                m.close()
            except Exception:
                pass
            try:
                m.logout()
            except Exception:
                pass


class MailPoller:
    """فحص دوري للبريد."""

    def __init__(self, make_fetcher, dest_dir, log, interval=300):
        self.make_fetcher = make_fetcher
        self.dest_dir = dest_dir
        self.log = log
        self.interval = max(60, int(interval))
        self._stop = threading.Event()
        self._thread = None
        self.last_run = None

    @property
    def running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.running:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        self.log(f"📧 فحص البريد كل {self.interval // 60} دقيقة.", "ok")

    def stop(self):
        self._stop.set()
        self.log("⏸ أُوقف فحص البريد.", "warn")

    def run_once(self) -> int:
        try:
            got = self.make_fetcher().fetch_pdfs(self.dest_dir, self.log)
            self.last_run = datetime.datetime.now()
            return len(got)
        except Exception as e:
            self.log(f"❌ البريد: {e}", "err")
            return 0

    def _loop(self):
        while not self._stop.is_set():
            self.run_once()
            self._stop.wait(self.interval)


# ══════════════════════════════════════════════════════════════════════
#  الواجهة الرسومية
# ══════════════════════════════════════════════════════════════════════

C_BG      = "#f1f5f9"
C_CARD    = "#ffffff"
C_PRIMARY = "#1d4ed8"
C_OK      = "#047857"
C_WARN    = "#b45309"
C_ERR     = "#b91c1c"
C_INK     = "#0f172a"
C_MUTED   = "#64748b"
C_TERM_BG = "#0f172a"
C_TERM_FG = "#e2e8f0"

F_UI    = ("Segoe UI", 10)
F_UI_B  = ("Segoe UI", 10, "bold")
F_TITLE = ("Segoe UI", 17, "bold")
F_MONO  = ("Consolas", 9)


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.cfg = load_config()

        root.title(f"{APP_NAME}  v{APP_VERSION}")
        root.geometry("980x720")
        root.minsize(860, 620)
        root.configure(bg=C_BG)

        self.provider_name = tk.StringVar(value=self.cfg.get("provider", "Gemini"))
        self.api_keys = {
            "Gemini":   tk.StringVar(value=self.cfg.get("gemini_key", "")),
            "Claude":   tk.StringVar(value=self.cfg.get("claude_key", "")),
            "Offline":  tk.StringVar(value=""),   # لا مفتاح مطلوب
            "OCR.space": tk.StringVar(value=self.cfg.get("ocrspace_key", "")),
        }
        self.model = tk.StringVar(value=self.cfg.get("model", GEMINI_MODELS[0]))
        self.excel_path = tk.StringVar(value=self.cfg.get("excel_path", DEFAULT_EXCEL))
        self.pdf_path = tk.StringVar()
        self.overwrite = tk.BooleanVar(value=self.cfg.get("overwrite", False))
        self.fallback = tk.BooleanVar(value=self.cfg.get("fallback", True))
        self.workspace_id = tk.StringVar(value=self.cfg.get("workspace_id", ""))
        self.status = tk.StringVar(value="جاهز")

        # ── المراقبة والأتمتة ──
        here = Path(__file__).resolve().parent
        self.watch_dir = tk.StringVar(
            value=self.cfg.get("watch_dir", str(here / "ملفات_واردة")))
        self.watch_on = tk.BooleanVar(value=self.cfg.get("watch_on", False))
        self.auto_save = tk.BooleanVar(value=self.cfg.get("auto_save", False))
        self.watch_every = tk.IntVar(value=self.cfg.get("watch_every", 6))

        # ── البريد ──
        self.mail_preset = tk.StringVar(value=self.cfg.get("mail_preset", "Gmail"))
        self.mail_host = tk.StringVar(value=self.cfg.get("mail_host", "imap.gmail.com"))
        self.mail_port = tk.IntVar(value=self.cfg.get("mail_port", 993))
        self.mail_user = tk.StringVar(value=self.cfg.get("mail_user", ""))
        self.mail_pass = tk.StringVar(value=self.cfg.get("mail_pass", ""))
        self.mail_box = tk.StringVar(value=self.cfg.get("mail_box", "INBOX"))
        self.mail_subject = tk.StringVar(value=self.cfg.get("mail_subject", ""))
        self.mail_sender = tk.StringVar(value=self.cfg.get("mail_sender", ""))
        self.mail_on = tk.BooleanVar(value=self.cfg.get("mail_on", False))
        self.mail_every = tk.IntVar(value=self.cfg.get("mail_every", 10))
        self.mail_mark = tk.BooleanVar(value=self.cfg.get("mail_mark", True))

        self.data: dict | None = None
        self.busy = False
        self.cancel = threading.Event()
        self._editor = None
        self.current_job: "Job | None" = None

        self.queue = JobQueue(on_change=self._queue_changed)
        self.watcher: "FolderWatcher | None" = None
        self.poller: "MailPoller | None" = None
        self._worker_stop = threading.Event()

        self._style()
        self._build()
        self._on_provider_change()

        if self.cfg:
            self.log("🔑 تم تحميل الإعدادات المحفوظة.", "ok")
        self.log(f"الإصدار {APP_VERSION}  ·  anthropic {_pkg_version('anthropic')}"
                 f"  ·  google-genai {_pkg_version('google-genai')}"
                 f"  ·  pymupdf {_pkg_version('pymupdf')}", "dim")
        self.log("جاهز للعمل. ابدأ من تبويب «الاستخراج».")

        # عامل الطابور يعمل طوال حياة البرنامج
        threading.Thread(target=self._worker_loop, daemon=True).start()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        if self.watch_on.get():
            self.root.after(600, self.toggle_watch)
        if self.mail_on.get():
            self.root.after(1200, self.toggle_mail)

    def _on_close(self):
        self._worker_stop.set()
        for s in (self.watcher, self.poller):
            if s:
                try:
                    s.stop()
                except Exception:
                    pass
        self.save_settings_silent()
        self.root.destroy()

    # ─────────────────────────────────────────────────────────────
    def _style(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TNotebook", background=C_BG, borderwidth=0)
        st.configure("TNotebook.Tab", font=F_UI_B, padding=(18, 9))
        st.map("TNotebook.Tab",
               background=[("selected", C_CARD), ("!selected", "#dbe3ec")],
               foreground=[("selected", C_PRIMARY), ("!selected", C_MUTED)])
        st.configure("TFrame", background=C_BG)
        st.configure("Card.TFrame", background=C_CARD, relief="flat")
        st.configure("TLabel", background=C_BG, font=F_UI, foreground=C_INK)
        st.configure("Card.TLabel", background=C_CARD, font=F_UI, foreground=C_INK)
        st.configure("Hint.TLabel", background=C_CARD, font=("Segoe UI", 9), foreground=C_MUTED)
        st.configure("TLabelframe", background=C_BG, font=F_UI_B, foreground=C_INK)
        st.configure("TLabelframe.Label", background=C_BG, font=F_UI_B, foreground=C_PRIMARY)
        st.configure("TButton", font=F_UI, padding=6)
        st.configure("TCheckbutton", background=C_BG, font=F_UI)
        st.configure("Treeview", font=F_UI, rowheight=27, fieldbackground=C_CARD)
        st.configure("Treeview.Heading", font=F_UI_B)
        st.configure("TProgressbar", troughcolor="#dbe3ec", background=C_PRIMARY, thickness=6)

    def _btn(self, parent, text, cmd, bg=C_PRIMARY, fg="white", **kw):
        return tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                         font=kw.pop("font", F_UI_B), relief="flat", cursor="hand2",
                         activebackground=bg, activeforeground=fg,
                         padx=kw.pop("padx", 16), pady=kw.pop("pady", 7),
                         bd=0, highlightthickness=0, **kw)

    # ─────────────────────────────────────────────────────────────
    def _build(self):
        head = tk.Frame(self.root, bg=C_PRIMARY)
        head.pack(fill="x")
        inner = tk.Frame(head, bg=C_PRIMARY)
        inner.pack(fill="x", padx=22, pady=13)
        tk.Label(inner, text=f"🚕  {APP_NAME}", font=F_TITLE,
                 bg=C_PRIMARY, fg="white").pack(anchor="e")
        tk.Label(inner, text=f"ولاية البيض — 32          الإصدار {APP_VERSION}",
                 font=("Segoe UI", 9), bg=C_PRIMARY, fg="#bfd3f7").pack(anchor="e")

        # شريط الحالة يُحجز أولاً وإلا ابتلعه الدفتر الممتد
        bar = tk.Frame(self.root, bg="#dbe3ec")
        bar.pack(fill="x", side="bottom")
        tk.Label(bar, textvariable=self.status, bg="#dbe3ec", fg=C_INK,
                 font=("Segoe UI", 9), anchor="e").pack(side="right", padx=14, pady=5)
        self.prov_lbl = tk.Label(bar, text="", bg="#dbe3ec", fg=C_MUTED,
                                 font=("Segoe UI", 9), anchor="w")
        self.prov_lbl.pack(side="left", padx=14, pady=5)

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(12, 6))
        self.tab_run = ttk.Frame(self.nb)
        self.tab_q   = ttk.Frame(self.nb)
        self.tab_rev = ttk.Frame(self.nb)
        self.tab_mail = ttk.Frame(self.nb)
        self.tab_set = ttk.Frame(self.nb)
        self.nb.add(self.tab_run,  text="📄  الاستخراج")
        self.nb.add(self.tab_q,    text="📥  الطابور والمراقبة")
        self.nb.add(self.tab_rev,  text="📋  مراجعة البيانات")
        self.nb.add(self.tab_mail, text="📧  البريد")
        self.nb.add(self.tab_set,  text="⚙️  الإعدادات")

        self._build_run()
        self._build_queue()
        self._build_review()
        self._build_mail()
        self._build_settings()

    # ── تبويب الاستخراج ──────────────────────────────────────────
    def _build_run(self):
        f = self.tab_run
        box = ttk.LabelFrame(f, text="  الملفات  ", padding=14)
        box.pack(fill="x", padx=14, pady=(14, 8))
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="ملف PDF للسائق:").grid(row=0, column=2, sticky="e", padx=(8, 6), pady=6)
        tk.Entry(box, textvariable=self.pdf_path, font=F_UI, relief="solid", bd=1
                 ).grid(row=0, column=1, sticky="ew", pady=6)
        self._btn(box, "اختر ملف…", self.browse_pdf, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).grid(row=0, column=0, padx=(6, 0))

        ttk.Label(box, text="ملف Excel:").grid(row=1, column=2, sticky="e", padx=(8, 6), pady=6)
        tk.Entry(box, textvariable=self.excel_path, font=F_UI, relief="solid", bd=1
                 ).grid(row=1, column=1, sticky="ew", pady=6)
        self._btn(box, "اختر ملف…", self.browse_excel, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).grid(row=1, column=0, padx=(6, 0))

        act = ttk.Frame(f)
        act.pack(fill="x", padx=14, pady=4)
        self.b_extract = self._btn(act, "🔍   استخراج البيانات", self.start_extract,
                                   font=("Segoe UI", 11, "bold"), padx=24, pady=10)
        self.b_extract.pack(side="right", padx=(8, 0))
        self.b_save = self._btn(act, "💾   حفظ في Excel", self.start_save, bg=C_OK,
                                font=("Segoe UI", 11, "bold"), padx=24, pady=10)
        self.b_save.pack(side="right", padx=8)
        self.b_save.config(state="disabled")
        self.b_stop = self._btn(act, "⏹  إيقاف", self.stop_work, bg="#be123c",
                                font=F_UI, padx=14, pady=10)
        self.b_stop.pack(side="right", padx=8)
        self.b_stop.config(state="disabled")
        self._btn(act, "🗑  مسح", self.clear, bg="#e2e8f0", fg=C_INK, font=F_UI).pack(side="right", padx=8)

        opts = ttk.Frame(act)
        opts.pack(side="left")
        ttk.Checkbutton(opts, text="تحديث الصف إن وُجد رقم الباب مسبقاً",
                        variable=self.overwrite).pack(anchor="w")
        ttk.Checkbutton(opts, text="جرّب نموذجاً بديلاً عند ازدحام الخادم",
                        variable=self.fallback).pack(anchor="w")

        self.pb = ttk.Progressbar(f, mode="determinate", maximum=100)
        self.pb.pack(fill="x", padx=14, pady=(6, 2))

        lf = ttk.LabelFrame(f, text="  سجلّ العمليات  ", padding=8)
        lf.pack(fill="both", expand=True, padx=14, pady=(6, 14))
        self.term = scrolledtext.ScrolledText(lf, font=F_MONO, bg=C_TERM_BG, fg=C_TERM_FG,
                                              relief="flat", wrap="word", height=12,
                                              insertbackground="white", state="disabled")
        self.term.pack(fill="both", expand=True)
        for tag, col in (("ok", "#4ade80"), ("warn", "#fbbf24"),
                         ("err", "#f87171"), ("info", "#60a5fa"), ("dim", "#94a3b8")):
            self.term.tag_config(tag, foreground=col)

    # ══ تبويب الطابور والمراقبة ══════════════════════════════════
    def _build_queue(self):
        f = self.tab_q

        ctl = ttk.LabelFrame(f, text="  المجلّد المراقَب  ", padding=12)
        ctl.pack(fill="x", padx=14, pady=(14, 8))
        ctl.columnconfigure(1, weight=1)

        ttk.Label(ctl, text="المجلّد:").grid(row=0, column=3, sticky="e", padx=(8, 6), pady=6)
        tk.Entry(ctl, textvariable=self.watch_dir, font=F_UI, relief="solid", bd=1
                 ).grid(row=0, column=1, columnspan=2, sticky="ew", pady=6)
        self._btn(ctl, "اختر…", self.browse_watch, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).grid(row=0, column=0, padx=(6, 0))

        row = ttk.Frame(ctl)
        row.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        self.b_watch = self._btn(row, "▶  تشغيل المراقبة", self.toggle_watch, bg=C_OK)
        self.b_watch.pack(side="right", padx=(0, 6))
        self._btn(row, "🔄  مسح الآن", self.scan_now, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).pack(side="right", padx=6)
        self._btn(row, "➕  إضافة ملفات…", self.add_files, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).pack(side="right", padx=6)
        self._btn(row, "📂  فتح المجلّد", self.open_watch, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).pack(side="right", padx=6)

        opt = ttk.Frame(ctl)
        opt.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        ttk.Checkbutton(opt, text="إدخال تلقائي للسجلّات التي تجتاز الفحص بلا تنبيه",
                        variable=self.auto_save).pack(side="right")
        ttk.Label(opt, text="   فحص كل:").pack(side="right", padx=(16, 2))
        tk.Spinbox(opt, from_=3, to=120, width=4, textvariable=self.watch_every,
                   font=F_UI).pack(side="right")
        ttk.Label(opt, text="ثانية").pack(side="right", padx=4)

        # لوحة العدّادات
        st = ttk.Frame(f)
        st.pack(fill="x", padx=14, pady=4)
        self.stat_lbls = {}
        for s in (ST_PENDING, ST_WORKING, ST_REVIEW, ST_SAVED, ST_SKIPPED, ST_FAILED):
            cell = tk.Frame(st, bg=C_CARD, highlightbackground="#cbd5e1",
                            highlightthickness=1)
            cell.pack(side="right", expand=True, fill="x", padx=3)
            tk.Label(cell, text=s, bg=C_CARD, fg=C_MUTED,
                     font=("Segoe UI", 8)).pack(pady=(5, 0))
            v = tk.Label(cell, text="0", bg=C_CARD, fg=ST_COLOR[s],
                         font=("Segoe UI", 15, "bold"))
            v.pack(pady=(0, 5))
            self.stat_lbls[s] = v

        # صفّ الأزرار يُحجز قبل الجدول الممتدّ وإلا ابتلعه
        bot = ttk.Frame(f)
        bot.pack(fill="x", side="bottom", padx=14, pady=(0, 14))
        self._btn(bot, "فتح للمراجعة", self._open_job, bg=C_PRIMARY, font=F_UI
                  ).pack(side="right")
        self._btn(bot, "🗑 تنظيف المنتهية", self.clear_done, bg="#e2e8f0",
                  fg=C_INK, font=F_UI).pack(side="right", padx=8)
        ttk.Label(bot, text="انقر نقراً مزدوجاً على أي سطر لفتحه في تبويب المراجعة.",
                  foreground=C_MUTED).pack(side="left")

        # الجدول
        wrap = ttk.Frame(f)
        wrap.pack(fill="both", expand=True, padx=14, pady=6)
        cols = ("status", "warn", "source", "name", "time")
        self.qtree = ttk.Treeview(wrap, columns=cols, show="headings", height=11)
        for c, t, w in (("time", "الوقت", 70), ("name", "الملف / السجلّ", 300),
                        ("source", "المصدر", 70), ("warn", "تنبيهات", 70),
                        ("status", "الحالة", 120)):
            self.qtree.heading(c, text=t, anchor="e")
            self.qtree.column(c, width=w, anchor="e",
                              stretch=(c == "name"))
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.qtree.yview)
        self.qtree.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        self.qtree.pack(side="left", fill="both", expand=True)
        for s, col in ST_COLOR.items():
            self.qtree.tag_configure(s, foreground=col)
        self.qtree.bind("<Double-1>", self._open_job)


    # ══ تبويب البريد ═════════════════════════════════════════════
    def _build_mail(self):
        f = self.tab_mail

        box = ttk.LabelFrame(f, text="  حساب البريد (IMAP)  ", padding=12)
        box.pack(fill="x", padx=14, pady=(14, 8))
        box.columnconfigure(1, weight=1)

        def row(r, label, var, show=None, width=None):
            ttk.Label(box, text=label).grid(row=r, column=2, sticky="e",
                                            padx=(8, 6), pady=5)
            e = tk.Entry(box, textvariable=var, font=F_UI, relief="solid", bd=1,
                         show=show or "")
            e.grid(row=r, column=1, sticky="ew", pady=5)
            return e

        ttk.Label(box, text="المزوّد:").grid(row=0, column=2, sticky="e", padx=(8, 6), pady=5)
        cb = ttk.Combobox(box, textvariable=self.mail_preset, font=F_UI, state="readonly",
                          values=list(IMAP_PRESETS))
        cb.grid(row=0, column=1, sticky="ew", pady=5)
        cb.bind("<<ComboboxSelected>>", self._mail_preset_changed)

        row(1, "الخادم:", self.mail_host)
        ttk.Label(box, text="المنفذ:").grid(row=2, column=2, sticky="e", padx=(8, 6), pady=5)
        tk.Spinbox(box, from_=1, to=65535, textvariable=self.mail_port, width=8,
                   font=F_UI).grid(row=2, column=1, sticky="w", pady=5)
        row(3, "البريد:", self.mail_user)
        self.e_mpass = row(4, "كلمة المرور:", self.mail_pass, show="•")
        self._btn(box, "👁", lambda: self.e_mpass.config(
            show="" if self.e_mpass.cget("show") else "•"),
            bg="#e2e8f0", fg=C_INK, font=F_UI, padx=10).grid(row=4, column=0, padx=(6, 0))
        self.mail_hint = ttk.Label(box, text="", foreground=C_MUTED, font=("Segoe UI", 9))
        self.mail_hint.grid(row=5, column=1, sticky="e", pady=(0, 6))

        flt = ttk.LabelFrame(f, text="  التصفية  ", padding=12)
        flt.pack(fill="x", padx=14, pady=8)
        flt.columnconfigure(1, weight=1)
        ttk.Label(flt, text="الصندوق:").grid(row=0, column=2, sticky="e", padx=(8, 6), pady=5)
        tk.Entry(flt, textvariable=self.mail_box, font=F_UI, relief="solid", bd=1
                 ).grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Label(flt, text="المُرسِل يحوي:").grid(row=1, column=2, sticky="e", padx=(8, 6), pady=5)
        tk.Entry(flt, textvariable=self.mail_sender, font=F_UI, relief="solid", bd=1
                 ).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(flt, text="الموضوع يحوي:").grid(row=2, column=2, sticky="e", padx=(8, 6), pady=5)
        tk.Entry(flt, textvariable=self.mail_subject, font=F_UI, relief="solid", bd=1
                 ).grid(row=2, column=1, sticky="ew", pady=5)
        o = ttk.Frame(flt)
        o.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Checkbutton(o, text="تعليم الرسالة كمقروءة بعد تنزيل مرفقاتها",
                        variable=self.mail_mark).pack(side="right")
        ttk.Label(o, text="   فحص كل:").pack(side="right", padx=(16, 2))
        tk.Spinbox(o, from_=1, to=240, width=4, textvariable=self.mail_every,
                   font=F_UI).pack(side="right")
        ttk.Label(o, text="دقيقة").pack(side="right", padx=4)

        act = ttk.Frame(f)
        act.pack(fill="x", padx=14, pady=10)
        self.b_mail = self._btn(act, "▶  تشغيل فحص البريد", self.toggle_mail, bg=C_OK,
                                padx=20, pady=9)
        self.b_mail.pack(side="right", padx=(0, 8))
        self._btn(act, "📥  فحص الآن", self.mail_now, padx=18, pady=9).pack(side="right", padx=8)
        self._btn(act, "🔌  اختبار الاتصال", self.mail_test, bg="#e2e8f0", fg=C_INK,
                  padx=18, pady=9).pack(side="right", padx=8)

        info = ttk.LabelFrame(f, text="  كيف يعمل  ", padding=12)
        info.pack(fill="both", expand=True, padx=14, pady=(4, 14))
        t = tk.Text(info, font=F_UI, relief="flat", bg=C_CARD, fg=C_INK, wrap="word")
        t.pack(fill="both", expand=True)
        t.insert("1.0",
            "يفحص البرنامج الرسائل غير المقروءة، ينزّل مرفقات PDF منها إلى المجلّد\n"
            "المراقَب، فتدخل خطّ المعالجة نفسه: استخراج ← فحص ← مراجعة أو إدخال.\n\n"
            "كلمة المرور:\n"
            "  • Gmail وOutlook يرفضان كلمة مرور الحساب العادية.\n"
            "    أنشئ «كلمة مرور التطبيقات» من إعدادات أمان حسابك واستعملها هنا.\n"
            "    Gmail: myaccount.google.com ← الأمان ← التحقّق بخطوتين ← كلمات مرور التطبيقات\n\n"
            "التصفية اختيارية. اتركها فارغة ليفحص كل الرسائل غير المقروءة،\n"
            "أو حدّد مُرسِلاً أو كلمة في الموضوع لتضييق النطاق.\n\n"
            "كلمة المرور تُحفظ في config.json بجانب البرنامج — لا تشارك هذا الملف.")
        t.config(state="disabled")

    # ── تبويب المراجعة ───────────────────────────────────────────
    def _build_review(self):
        f = self.tab_rev
        top = ttk.Frame(f)
        top.pack(fill="x", padx=14, pady=(14, 6))
        ttk.Label(top, text="انقر نقراً مزدوجاً على أي قيمة لتصحيحها قبل الحفظ.",
                  foreground=C_MUTED).pack(side="right")
        self._btn(top, "↻ إعادة الفحص", self.revalidate, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).pack(side="left")
        self._btn(top, "🔧 إصلاح واشتقاق", lambda: self.revalidate(fix=True),
                  bg="#fef3c7", fg=C_INK, font=F_UI).pack(side="left", padx=8)

        # لوحة الفحص تُحجز قبل الجدول الممتد حتى تبقى ظاهرة دائماً
        vf = ttk.LabelFrame(f, text="  نتيجة الفحص  ", padding=8)
        vf.pack(fill="x", side="bottom", padx=14, pady=(6, 14))
        self.vbox = tk.Text(vf, height=5, font=F_UI, relief="flat",
                            bg=C_CARD, fg=C_INK, wrap="word", state="disabled")
        self.vbox.pack(fill="x")
        self.vbox.tag_config("ok", foreground=C_OK)
        self.vbox.tag_config("warn", foreground=C_WARN)

        wrap = ttk.Frame(f)
        wrap.pack(fill="both", expand=True, padx=14, pady=4)
        cols = ("value", "field")
        self.tree = ttk.Treeview(wrap, columns=cols, show="tree headings", height=12)
        self.tree.heading("#0", text="القسم / الحقل", anchor="e")
        self.tree.heading("value", text="القيمة", anchor="e")
        self.tree.heading("field", text="المفتاح", anchor="e")
        self.tree.column("#0", width=290, minwidth=180, anchor="e", stretch=False)
        self.tree.column("value", width=380, minwidth=160, anchor="e", stretch=True)
        self.tree.column("field", width=175, minwidth=110, anchor="e", stretch=False)
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("group", font=F_UI_B, background="#e8eef7")
        self.tree.tag_configure("empty", foreground="#b0b8c4")
        self.tree.tag_configure("bad", foreground=C_ERR)
        self.tree.bind("<Double-1>", self._edit_cell)

    # ── تبويب الإعدادات ──────────────────────────────────────────
    def _build_settings(self):
        f = self.tab_set
        box = ttk.LabelFrame(f, text="  مزوّد الذكاء الاصطناعي  ", padding=14)
        box.pack(fill="x", padx=14, pady=(14, 8))
        box.columnconfigure(1, weight=1)

        ttk.Label(box, text="المزوّد:").grid(row=0, column=2, sticky="e", padx=(8, 6), pady=7)
        pf = ttk.Frame(box)
        pf.grid(row=0, column=1, sticky="w", pady=7)
        for name in PROVIDERS:
            tk.Radiobutton(pf, text=name, variable=self.provider_name, value=name,
                           command=self._on_provider_change, bg=C_BG, font=F_UI,
                           activebackground=C_BG, selectcolor="#cfe0ff"
                           ).pack(side="left", padx=(0, 16))

        ttk.Label(box, text="مفتاح API:").grid(row=1, column=2, sticky="e", padx=(8, 6), pady=7)
        self.e_key = tk.Entry(box, font=("Consolas", 10), show="•", relief="solid", bd=1)
        self.e_key.grid(row=1, column=1, sticky="ew", pady=7)
        kb = ttk.Frame(box)
        kb.grid(row=1, column=0, padx=(6, 0))
        self._btn(kb, "👁", self._toggle_key, bg="#e2e8f0", fg=C_INK, font=F_UI, padx=10
                  ).pack(side="left", padx=2)

        self.key_hint = ttk.Label(box, text="", foreground=C_MUTED, font=("Segoe UI", 9))
        self.key_hint.grid(row=2, column=1, sticky="e", pady=(0, 6))

        # مساحة العمل — لمفاتيح Claude غير المرتبطة بمساحة
        self.ws_lbl = ttk.Label(box, text="مساحة العمل:")
        self.ws_lbl.grid(row=4, column=2, sticky="e", padx=(8, 6), pady=7)
        self.e_ws = tk.Entry(box, textvariable=self.workspace_id,
                             font=("Consolas", 10), relief="solid", bd=1)
        self.e_ws.grid(row=4, column=1, sticky="ew", pady=7)
        self.ws_hint = ttk.Label(box, text="اختياري — يلزم فقط إن رفض المفتاح الطلب",
                                 foreground=C_MUTED, font=("Segoe UI", 9))
        self.ws_hint.grid(row=5, column=1, sticky="e", pady=(0, 6))

        ttk.Label(box, text="النموذج:").grid(row=3, column=2, sticky="e", padx=(8, 6), pady=7)
        self.cb_model = ttk.Combobox(box, textvariable=self.model, font=F_UI, state="normal")
        self.cb_model.grid(row=3, column=1, sticky="ew", pady=7)
        self._btn(box, "↻ جلب النماذج", self.fetch_models, bg="#e2e8f0", fg=C_INK, font=F_UI
                  ).grid(row=3, column=0, padx=(6, 0))

        act = ttk.Frame(f)
        act.pack(fill="x", padx=14, pady=10)
        self._btn(act, "💾  حفظ الإعدادات", self.save_settings, bg=C_OK,
                  padx=20, pady=9).pack(side="right", padx=(8, 0))
        self._btn(act, "🔌  اختبار الاتصال", self.test_connection,
                  padx=20, pady=9).pack(side="right", padx=8)
        self._btn(act, "⬇  تثبيت اللغة العربية", self._install_tessdata,
                  bg="#7c3aed", fg="white", padx=16, pady=9).pack(side="right", padx=8)

        info = ttk.LabelFrame(f, text="  معلومات  ", padding=14)
        info.pack(fill="both", expand=True, padx=14, pady=(4, 14))
        txt = tk.Text(info, font=F_UI, relief="flat", bg=C_CARD, fg=C_INK,
                      wrap="word", height=10)
        txt.pack(fill="both", expand=True)
        txt.insert("1.0",
            "الحصول على المفاتيح:\n"
            "  • Gemini  (طبقة مجانية سخية)  →  aistudio.google.com/apikey\n"
            "  • Claude  (مدفوع بالرصيد)     →  console.anthropic.com\n"
            "  • OCR.space (مجاني: 500 طلب يومياً، 1 م.ب للصفحة)  →  ocr.space/ocrapi/freekey\n"
            "    (يمكن أيضاً وضع المفتاح في OCR_SPACE_API_KEY أعلى ocr_space.py)\n\n"
            "ملف الإعدادات config.json يُحفظ بجانب البرنامج ويحتوي المفاتيح،\n"
            "لذا لا تشارك هذا الملف مع أحد.\n\n"
            "قبل كل كتابة في Excel يُنشئ البرنامج نسخة احتياطية تلقائية\n"
            "باسم  TAXI 32000_backup_<التاريخ>.xlsx  في نفس المجلد.\n\n"
            "الأوراق المستهدفة:  Operateur · Chauffeur · Véhicule · Licence Modjahid\n"
            "يبدأ البحث عن أول صف فارغ من الصف 3 (الصفان 1 و2 رؤوس أعمدة).")
        txt.config(state="disabled")

    # ─────────────────────────────────────────────────────────────
    #  أدوات مساعدة
    # ─────────────────────────────────────────────────────────────
    def log(self, msg: str, tag: str = ""):
        def _w():
            self.term.config(state="normal")
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self.term.insert("end", f"[{ts}] ", "dim")
            self.term.insert("end", msg + "\n", tag)
            self.term.see("end")
            self.term.config(state="disabled")
        self.root.after(0, _w)

    def set_status(self, s: str):
        self.root.after(0, lambda: self.status.set(s))

    def set_progress(self, v: float):
        self.root.after(0, lambda: self.pb.config(value=v))

    def _toggle_key(self):
        self.e_key.config(show="" if self.e_key.cget("show") else "•")

    def _on_provider_change(self):
        p = self.provider_name.get()
        # ── Offline: لا مفتاح ولا نموذج سحابي ──────────────────────────────
        if p == "Offline":
            self.e_key.delete(0, "end")
            self.e_key.config(state="disabled")
            self.cb_model["values"] = OFFLINE_MODELS
            self.model.set(OFFLINE_MODELS[0])
            self.key_hint.config(text="لا يحتاج مفتاح API — OCR محلي بالكامل ✓")
            for w in (self.ws_lbl, self.e_ws, self.ws_hint):
                w.grid_remove()
            self.prov_lbl.config(text="المزوّد: Offline (Tesseract)  •  بدون إنترنت")
            return
        self.e_key.config(state="normal")
        self.e_key.delete(0, "end")
        self.e_key.insert(0, self.api_keys.get(p, tk.StringVar()).get())
        # ── OCR.space: مفتاح مجاني، ومحرّكان بدل النماذج ───────────────────
        if p == "OCR.space":
            self.cb_model["values"] = OCRSPACE_MODELS
            if self.model.get() not in OCRSPACE_MODELS:
                self.model.set(OCRSPACE_MODELS[0])
            self.key_hint.config(text="مفتاح مجاني من ocr.space/ocrapi/freekey — "
                                      "حتى 1 ميغابايت للصفحة و500 طلب يومياً")
            for w in (self.ws_lbl, self.e_ws, self.ws_hint):
                w.grid_remove()
            self.prov_lbl.config(text=f"المزوّد: OCR.space  •  {self.model.get()}")
            return
        models = GEMINI_MODELS if p == "Gemini" else CLAUDE_MODELS
        self.cb_model["values"] = models
        if self.model.get() not in models:
            self.model.set(models[0])
        self.key_hint.config(
            text="احصل على مفتاح مجاني من aistudio.google.com/apikey" if p == "Gemini"
            else "يتطلب رصيداً مدفوعاً — console.anthropic.com")
        # خانة مساحة العمل تخصّ Claude وحده
        for w in (self.ws_lbl, self.e_ws, self.ws_hint):
            (w.grid if p == "Claude" else w.grid_remove)()
        self.prov_lbl.config(text=f"المزوّد: {p}  •  النموذج: {self.model.get()}")

    def _sync_key(self):
        p = self.provider_name.get()
        if p != "Offline":
            self.api_keys[p].set(self.e_key.get().strip())

    def _provider(self):
        p = self.provider_name.get()
        if p == "Offline":
            prov = OfflineProvider()
            prov.model = self.model.get().strip() or OFFLINE_MODELS[0]
            return prov
        self._sync_key()
        key = self.api_keys.get(p, tk.StringVar()).get().strip()
        if p == "OCR.space":
            # المفتاح من الإعدادات، وإلا من OCR_SPACE_API_KEY في ocr_space.py
            return OcrSpaceProvider(key, self.model.get().strip())
        if not key:
            raise ProviderError(f"مفتاح {p} غير مُدخل — اذهب إلى تبويب الإعدادات.")
        if p == "Claude":
            return ClaudeProvider(key, self.model.get().strip(),
                                  self.workspace_id.get().strip())
        return PROVIDERS[p](key, self.model.get().strip())

    def browse_pdf(self):
        p = filedialog.askopenfilename(title="اختر ملف PDF",
                                       filetypes=[("PDF", "*.pdf"), ("كل الملفات", "*.*")])
        if p:
            self.pdf_path.set(p)
            self.log(f"📂 {Path(p).name}", "info")

    def browse_excel(self):
        p = filedialog.askopenfilename(title="اختر ملف Excel",
                                       filetypes=[("Excel", "*.xlsx"), ("كل الملفات", "*.*")])
        if p:
            self.excel_path.set(p)

    def clear(self):
        self.term.config(state="normal")
        self.term.delete("1.0", "end")
        self.term.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self._set_validation([])
        self.data = None
        self.b_save.config(state="disabled")
        self.pb.config(value=0)
        self.set_status("جاهز")

    def save_settings(self):
        self._sync_key()
        ok = save_config({
            "provider":   self.provider_name.get(),
            "gemini_key": self.api_keys["Gemini"].get().strip(),
            "claude_key": self.api_keys["Claude"].get().strip(),
            "ocrspace_key": self.api_keys["OCR.space"].get().strip(),
            "model":      self.model.get().strip(),
            "excel_path": self.excel_path.get().strip(),
            "overwrite":  bool(self.overwrite.get()),
            "fallback":   bool(self.fallback.get()),
            "workspace_id": self.workspace_id.get().strip(),
            "watch_dir":  self.watch_dir.get().strip(),
            "watch_on":   bool(self.watch_on.get()),
            "watch_every": int(self.watch_every.get()),
            "auto_save":  bool(self.auto_save.get()),
            "mail_preset": self.mail_preset.get(),
            "mail_host":  self.mail_host.get().strip(),
            "mail_port":  int(self.mail_port.get()),
            "mail_user":  self.mail_user.get().strip(),
            "mail_pass":  self.mail_pass.get(),
            "mail_box":   self.mail_box.get().strip(),
            "mail_subject": self.mail_subject.get().strip(),
            "mail_sender": self.mail_sender.get().strip(),
            "mail_on":    bool(self.mail_on.get()),
            "mail_every": int(self.mail_every.get()),
            "mail_mark":  bool(self.mail_mark.get()),
        })
        self._on_provider_change()
        if ok:
            self.log("💾 تم حفظ الإعدادات.", "ok")
            messagebox.showinfo("تم", "حُفظت الإعدادات.\nستُحمَّل تلقائياً عند التشغيل القادم.")
        else:
            messagebox.showerror("خطأ", "تعذّر حفظ ملف الإعدادات.")

    # ── اختبار الاتصال وجلب النماذج ──────────────────────────────

    def _install_tessdata(self):
        """تنزيل ara.traineddata و fra.traineddata من GitHub في خيط خلفي.
        الفرنسية لازمة لقراءة الأسماء اللاتينية وبيانات البطاقة الرمادية."""
        import urllib.request, os
        # نفس المجلد الذي يبحث فيه البرنامج أولاً (مسار ASCII يقبله Tesseract)
        dest_dir = Path(_find_tessdata() or TESSDATA_DIRS[0])
        dest_dir.mkdir(parents=True, exist_ok=True)
        urls = {
            "ara": "https://github.com/tesseract-ocr/tessdata/raw/main/ara.traineddata",
            "fra": "https://github.com/tesseract-ocr/tessdata_fast/raw/main/fra.traineddata",
        }
        todo = [l for l in urls if not (dest_dir / f"{l}.traineddata").exists()]
        if not todo:
            messagebox.showinfo("اللغات موجودة",
                                f"ara و fra موجودتان بالفعل في:\n{dest_dir}")
            return
        self.log(f"⬇ جارٍ تنزيل: {', '.join(todo)} ← {dest_dir}", "info")
        def job():
            try:
                for l in todo:
                    dest = dest_dir / f"{l}.traineddata"
                    urllib.request.urlretrieve(urls[l], str(dest))
                    self.log(f"✅ تم حفظ: {dest}", "ok")
                messagebox.showinfo("تمّ التثبيت",
                    f"تم تنزيل {', '.join(todo)} بنجاح.\n{dest_dir}\n\nاختر Offline وابدأ الاستخراج.")
            except Exception as e:
                self.log(f"❌ فشل التنزيل: {e}", "err")
                messagebox.showerror("خطأ", str(e))
        threading.Thread(target=job, daemon=True).start()

    def test_connection(self):
        def job():
            try:
                self.set_status("جارٍ اختبار الاتصال…")
                prov = self._provider()
                msg = prov.check()
                self.log(f"🔌 {msg}", "ok")
                # جلب القائمة الحقيقية فوراً حتى تظهر أسماء النماذج الفعلية
                try:
                    models = prov.list_models()
                    if models:
                        self.root.after(0, lambda: self.cb_model.config(values=models))
                        self.log(f"↻ النماذج المتاحة ({len(models)}): " +
                                 "، ".join(models[:4]) + ("…" if len(models) > 4 else ""), "info")
                        if self.model.get() not in models:
                            self.root.after(0, lambda: self.model.set(models[0]))
                            self.log(f"↪ النموذج المختار غير متاح — تم اختيار {models[0]}", "warn")
                except Exception:
                    pass
                self.set_status("الاتصال سليم")
                messagebox.showinfo("نجح", msg)
            except Exception as e:
                self.log(f"❌ فشل الاتصال: {e}", "err")
                self.set_status("فشل الاتصال")
                messagebox.showerror("فشل الاتصال", str(e))
        threading.Thread(target=job, daemon=True).start()

    def fetch_models(self):
        def job():
            try:
                self.set_status("جارٍ جلب النماذج…")
                models = self._provider().list_models()
                if models:
                    self.root.after(0, lambda: self.cb_model.config(values=models))
                    self.log(f"↻ {len(models)} نموذج متاح.", "ok")
                self.set_status("جاهز")
            except Exception as e:
                self.log(f"❌ تعذّر جلب النماذج: {e}", "err")
                self.set_status("جاهز")
        threading.Thread(target=job, daemon=True).start()

    # ── الاستخراج ────────────────────────────────────────────────
    def start_extract(self):
        if self.busy:
            return
        pdf = self.pdf_path.get().strip()
        if not pdf:
            messagebox.showwarning("تنبيه", "اختر ملف PDF أولاً.")
            return
        if not Path(pdf).is_file():
            messagebox.showerror("خطأ", "ملف PDF غير موجود.")
            return
        self.busy = True
        self.cancel.clear()
        self.b_extract.config(state="disabled")
        self.b_save.config(state="disabled")
        self.b_stop.config(state="normal")
        self.pb.config(value=0)
        threading.Thread(target=self._extract_job, args=(pdf,), daemon=True).start()

    def stop_work(self):
        self.cancel.set()
        self.log("⏹ طلب الإيقاف — سيتوقف بعد الخطوة الحالية.", "warn")
        self.set_status("جارٍ الإيقاف…")

    def _extract_job(self, pdf: str):
        """المسار اليدوي — يستعمل نفس محرّك الاستخراج المشترك."""
        try:
            data = self._extract_file(pdf)
            self.data = data
            self.current_job = None
            self.log("✅ تم الاستخراج بنجاح.", "ok")
            self.root.after(0, lambda: self._populate(data))
            self.save_settings_silent()
            self.set_status("تم الاستخراج — راجع البيانات ثم احفظ")
            self.root.after(0, lambda: self.b_save.config(state="normal"))
            self.root.after(0, lambda: self.nb.select(self.tab_rev))
        except Cancelled:
            self.log("⏹ أُوقفت العملية بناءً على طلبك.", "warn")
            self.set_status("مُوقفة")
            self.set_progress(0)
        except json.JSONDecodeError as e:
            self.log(f"❌ رد النموذج ليس JSON صالحاً: {e}", "err")
            self.set_status("فشل تحليل الرد")
        except Exception as e:
            self.log(f"❌ {e}", "err")
            if is_retryable(e):
                self.log("💡 الخادم مزدحم. جرّب بعد قليل أو فعّل «نموذج بديل».", "warn")
            self.set_status("حدث خطأ")
        finally:
            self.busy = False
            self.root.after(0, lambda: self.b_extract.config(state="normal"))
            self.root.after(0, lambda: self.b_stop.config(state="disabled"))

    def save_settings_silent(self):
        self._sync_key()
        save_config({
            "provider":   self.provider_name.get(),
            "gemini_key": self.api_keys["Gemini"].get().strip(),
            "claude_key": self.api_keys["Claude"].get().strip(),
            "ocrspace_key": self.api_keys["OCR.space"].get().strip(),
            "model":      self.model.get().strip(),
            "excel_path": self.excel_path.get().strip(),
            "overwrite":  bool(self.overwrite.get()),
            "fallback":   bool(self.fallback.get()),
            "workspace_id": self.workspace_id.get().strip(),
            "watch_dir":  self.watch_dir.get().strip(),
            "watch_on":   bool(self.watch_on.get()),
            "watch_every": int(self.watch_every.get()),
            "auto_save":  bool(self.auto_save.get()),
            "mail_preset": self.mail_preset.get(),
            "mail_host":  self.mail_host.get().strip(),
            "mail_port":  int(self.mail_port.get()),
            "mail_user":  self.mail_user.get().strip(),
            "mail_pass":  self.mail_pass.get(),
            "mail_box":   self.mail_box.get().strip(),
            "mail_subject": self.mail_subject.get().strip(),
            "mail_sender": self.mail_sender.get().strip(),
            "mail_on":    bool(self.mail_on.get()),
            "mail_every": int(self.mail_every.get()),
            "mail_mark":  bool(self.mail_mark.get()),
        })

    # ── جدول المراجعة ────────────────────────────────────────────
    def _populate(self, data: dict):
        self.tree.delete(*self.tree.get_children())
        for gkey, grp in SCHEMA.items():
            vals = data.get(gkey) or {}
            filled = sum(1 for f, _, _ in grp["fields"]
                         if vals.get(f) not in (None, "", "null"))
            node = self.tree.insert(
                "", "end",
                text=f"{grp['icon']}  {grp['title']}   ({filled}/{len(grp['fields'])})",
                values=("", gkey), open=True, tags=("group",))
            for fkey, flabel, _ in grp["fields"]:
                v = vals.get(fkey)
                shown = "" if v in (None, "null") else str(v)
                self.tree.insert(node, "end", text=f"   {flabel}",
                                 values=(shown, fkey),
                                 tags=() if shown else ("empty",))
        notes = data.get("notes")
        if notes:
            n = self.tree.insert("", "end", text="📝  ملاحظات", values=("", "notes"),
                                 open=True, tags=("group",))
            self.tree.insert(n, "end", text="   نص", values=(str(notes), "notes"))
        self.revalidate()

    def _edit_cell(self, event):
        if self._editor:
            self._editor.destroy()
            self._editor = None
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item or col != "#1":
            return
        parent = self.tree.parent(item)
        if not parent:
            return
        x, y, w, h = self.tree.bbox(item, col)
        cur = self.tree.set(item, "value")
        e = tk.Entry(self.tree, font=F_UI, relief="solid", bd=1, justify="right")
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, cur)
        e.focus_set()
        e.select_range(0, "end")
        self._editor = e

        def commit(_=None):
            new = e.get().strip()
            self.tree.set(item, "value", new)
            self.tree.item(item, tags=() if new else ("empty",))
            gkey = self.tree.set(parent, "field")
            fkey = self.tree.set(item, "field")
            if self.data is not None and gkey in self.data:
                self.data.setdefault(gkey, {})[fkey] = new or None
            elif self.data is not None and gkey == "notes":
                self.data["notes"] = new
            e.destroy()
            self._editor = None
            self.revalidate()

        def cancel(_=None):
            e.destroy()
            self._editor = None

        e.bind("<Return>", commit)
        e.bind("<FocusOut>", commit)
        e.bind("<Escape>", cancel)

    def revalidate(self, fix=False):
        if not self.data:
            self._set_validation([])
            return
        if fix:
            fx = normalize(self.data, log=lambda m: self.log(m, "warn"))
            fx += enrich(self.data, log=lambda m: self.log(m, "dim"))
            if fx:
                self.log(f"🔧 أُصلح {len(fx)} حقلاً.", "info")
                self._populate(self.data)
                return
            self.log("لا شيء يحتاج إصلاحاً تلقائياً.", "dim")
        self._set_validation(validate(self.data))

    def _set_validation(self, warns: list[str]):
        self.vbox.config(state="normal")
        self.vbox.delete("1.0", "end")
        if not self.data:
            self.vbox.insert("end", "لا توجد بيانات بعد.")
        elif not warns:
            self.vbox.insert("end", "✅  اجتازت البيانات كل الفحوصات.", "ok")
        else:
            self.vbox.insert("end", f"⚠️  {len(warns)} تنبيه:\n", "warn")
            for w in warns:
                self.vbox.insert("end", f"   • {w}\n", "warn")
        self.vbox.config(state="disabled")

    # ══ الطابور: العرض والتحكّم ══════════════════════════════════
    def _queue_changed(self):
        try:
            self.root.after(0, self._refresh_queue)
        except Exception:
            pass

    def _refresh_queue(self):
        counts = self.queue.counts()
        for s, lbl in self.stat_lbls.items():
            lbl.config(text=str(counts.get(s, 0)))
        sel = set(self.qtree.selection())
        self.qtree.delete(*self.qtree.get_children())
        for j in reversed(self.queue.all()):
            iid = str(j.id)
            self.qtree.insert(
                "", "end", iid=iid, tags=(j.status,),
                values=(j.status,
                        str(len(j.warnings)) if j.warnings else "—",
                        SRC_LABEL.get(j.source, j.source),
                        j.label(),
                        j.created.strftime("%H:%M:%S")))
        for i in sel:
            if self.qtree.exists(i):
                self.qtree.selection_add(i)
        act = counts.get(ST_PENDING, 0) + counts.get(ST_WORKING, 0)
        if act:
            self.set_status(f"الطابور: {act} قيد الانتظار/المعالجة")

    def _open_job(self, _=None):
        sel = self.qtree.selection()
        if not sel:
            return
        job = self.queue.get(int(sel[0]))
        if not job:
            return
        if not job.data:
            messagebox.showinfo("لا بيانات", "هذا الملف لم يُستخرج بعد أو فشل استخراجه."
                                + (f"\n\n{job.error}" if job.error else ""))
            return
        self.current_job = job
        self.data = job.data
        self._populate(job.data)
        self.b_save.config(state="normal")
        self.nb.select(self.tab_rev)

    def clear_done(self):
        self.queue.clear_done()
        self.log("🗑 نُظّفت السجلّات المنتهية من الطابور.", "dim")

    def browse_watch(self):
        d = filedialog.askdirectory(title="اختر المجلّد المراقَب")
        if d:
            self.watch_dir.set(d)

    def open_watch(self):
        p = Path(self.watch_dir.get())
        p.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(p))
            elif sys.platform == "darwin":
                os.system(f'open "{p}"')
            else:
                os.system(f'xdg-open "{p}"')
        except Exception as e:
            messagebox.showinfo("المجلّد", f"{p}\n\n{e}")

    def add_files(self):
        paths = filedialog.askopenfilenames(title="اختر ملفات PDF",
                                            filetypes=[("PDF", "*.pdf")])
        n = sum(1 for p in paths if self.queue.add(p, "manual"))
        if n:
            self.log(f"➕ أُضيف {n} ملف إلى الطابور.", "info")
            self.nb.select(self.tab_q)
        elif paths:
            self.log("كل الملفات المختارة موجودة في الطابور.", "dim")

    def _ensure_watcher(self):
        d = self.watch_dir.get().strip()
        if not self.watcher or str(self.watcher.folder) != d:
            self.watcher = FolderWatcher(d, self.queue, self.log,
                                         max(3, self.watch_every.get()))
        self.watcher.interval = max(3, self.watch_every.get())
        return self.watcher

    def scan_now(self):
        w = self._ensure_watcher()
        w.folder.mkdir(parents=True, exist_ok=True)
        # مسحة يدوية: لا ننتظر استقرار الحجم
        n = 0
        for p in sorted(w.folder.glob("*.pdf")):
            if p.is_file() and self.queue.add(p, "folder"):
                n += 1
        self.log(f"🔄 المسح اليدوي: {n} ملف جديد." if n else "🔄 لا ملفات جديدة.",
                 "info" if n else "dim")
        if n:
            self.nb.select(self.tab_q)

    def toggle_watch(self):
        w = self._ensure_watcher()
        if w.running:
            w.stop()
            self.watch_on.set(False)
            self.b_watch.config(text="▶  تشغيل المراقبة", bg=C_OK)
        else:
            w.start()
            self.watch_on.set(True)
            self.b_watch.config(text="⏸  إيقاف المراقبة", bg="#be123c")
        self.save_settings_silent()

    # ══ البريد ═══════════════════════════════════════════════════
    def _mail_preset_changed(self, _=None):
        host, port = IMAP_PRESETS.get(self.mail_preset.get(), ("", 993))
        if host:
            self.mail_host.set(host)
            self.mail_port.set(port)
        hints = {"Gmail": "استعمل «كلمة مرور التطبيقات» لا كلمة مرور الحساب.",
                 "Outlook/Hotmail": "قد تحتاج كلمة مرور تطبيقات إن فعّلت التحقّق بخطوتين.",
                 "Yahoo": "يتطلب «كلمة مرور التطبيقات» من إعدادات الحساب."}
        self.mail_hint.config(text=hints.get(self.mail_preset.get(), ""))

    def _make_fetcher(self):
        return MailFetcher(self.mail_host.get().strip(), self.mail_port.get(),
                           self.mail_user.get().strip(), self.mail_pass.get(),
                           self.mail_box.get().strip() or "INBOX",
                           self.mail_subject.get(), self.mail_sender.get(),
                           bool(self.mail_mark.get()))

    def mail_test(self):
        def job():
            try:
                self.set_status("اختبار البريد…")
                msg = self._make_fetcher().check()
                self.log(f"📧 {msg}", "ok")
                self.set_status("البريد متصل")
                self.root.after(0, lambda: messagebox.showinfo("نجح", msg))
            except Exception as e:
                self.log(f"❌ البريد: {e}", "err")
                self.set_status("فشل الاتصال بالبريد")
                self.root.after(0, lambda: messagebox.showerror("فشل الاتصال", str(e)))
        threading.Thread(target=job, daemon=True).start()

    def _mail_dest(self):
        d = Path(self.watch_dir.get().strip())
        d.mkdir(parents=True, exist_ok=True)
        return d

    def mail_now(self):
        def job():
            try:
                self.set_status("فحص البريد…")
                got = self._make_fetcher().fetch_pdfs(self._mail_dest(), self.log)
                for path, note in got:
                    self.queue.add(path, "email", note)
                self.log(f"📧 أُضيف {len(got)} مرفق إلى الطابور." if got
                         else "📧 لا مرفقات جديدة.", "ok" if got else "dim")
                self.set_status("جاهز")
                if got:
                    self.root.after(0, lambda: self.nb.select(self.tab_q))
            except Exception as e:
                self.log(f"❌ البريد: {e}", "err")
                self.set_status("فشل فحص البريد")
                self.root.after(0, lambda: messagebox.showerror("خطأ", str(e)))
        threading.Thread(target=job, daemon=True).start()

    def toggle_mail(self):
        if self.poller and self.poller.running:
            self.poller.stop()
            self.mail_on.set(False)
            self.b_mail.config(text="▶  تشغيل فحص البريد", bg=C_OK)
        else:
            if not self.mail_user.get().strip() or not self.mail_pass.get():
                messagebox.showwarning("تنبيه", "أدخل بريدك وكلمة المرور أولاً.")
                return

            def fetch_and_queue():
                f = self._make_fetcher()
                orig = f.fetch_pdfs

                def wrapped(dest, log, limit=25):
                    got = orig(dest, log, limit)
                    for path, note in got:
                        self.queue.add(path, "email", note)
                    return got
                f.fetch_pdfs = wrapped
                return f

            self.poller = MailPoller(fetch_and_queue, self._mail_dest(), self.log,
                                     max(60, self.mail_every.get() * 60))
            self.poller.start()
            self.mail_on.set(True)
            self.b_mail.config(text="⏸  إيقاف فحص البريد", bg="#be123c")
        self.save_settings_silent()

    # ══ عامل الطابور ═════════════════════════════════════════════
    def _worker_loop(self):
        """يعالج ملفات الطابور واحداً تلو الآخر دون تجميد الواجهة."""
        while not self._worker_stop.is_set():
            job = self.queue.next_pending()
            if job is None:
                self._worker_stop.wait(1.5)
                continue
            if self.busy:                      # لا نزاحم استخراجاً يدوياً
                self.queue.update(job, status=ST_PENDING)
                self._worker_stop.wait(2)
                continue
            self._process(job)

    def _process(self, job):
        self.busy = True
        try:
            self.log(f"▶ [{SRC_LABEL.get(job.source)}] {job.name}", "info")
            self.set_status(f"معالجة {job.name}…")
            data = self._extract_file(str(job.path))
            warns = validate(data)
            self.queue.update(job, data=data, warnings=warns)

            auto = bool(self.auto_save.get())
            if auto and not warns:
                s = write_to_excel(self.excel_path.get().strip(), data, False,
                                   lambda m: self.log("   " + m, "dim"))
                added = len(s["added"]) + len(s["updated"])
                if added:
                    self.queue.update(job, status=ST_SAVED,
                                      summary=f"{added} ورقة")
                    self.log(f"✅ أُدخل تلقائياً: {job.label()}", "ok")
                    self._file_away(job, "تم_إدخالها")
                else:
                    self.queue.update(job, status=ST_SKIPPED)
                    self.log(f"⏭️ مكرّر: {job.label()}", "warn")
                    self._file_away(job, "تم_إدخالها")
            else:
                self.queue.update(job, status=ST_REVIEW)
                why = f"{len(warns)} تنبيه" if warns else "الإدخال التلقائي مغلق"
                self.log(f"📋 بانتظار المراجعة ({why}): {job.label()}", "warn")
                if warns:
                    for w in warns[:3]:
                        self.log(f"     • {w}", "dim")
        except Cancelled:
            self.queue.update(job, status=ST_PENDING)
        except Exception as e:
            self.queue.update(job, status=ST_FAILED, error=str(e))
            self.log(f"❌ {job.name}: {e}", "err")
        finally:
            self.busy = False
            self.set_status("جاهز")

    def _file_away(self, job, subdir):
        """ينقل الملف المعالَج إلى مجلّد فرعي — فقط إن كان داخل المجلّد المراقَب."""
        try:
            watch = Path(self.watch_dir.get().strip()).resolve()
            if watch not in job.path.resolve().parents:
                return
            dest = watch / subdir
            dest.mkdir(exist_ok=True)
            target = dest / job.path.name
            i = 1
            while target.exists():
                target = dest / f"{job.path.stem}_{i}{job.path.suffix}"
                i += 1
            shutil.move(str(job.path), str(target))
            self.queue.update(job, path=target)
        except Exception as e:
            self.log(f"⚠️ تعذّر نقل {job.name}: {e}", "dim")

    # ══ الاستخراج المشترك (يدوي وآلي) ════════════════════════════
    def _extract_file(self, pdf: str) -> dict:
        """يستخرج ملفاً واحداً ويعيد dict — يرفع استثناءً عند الفشل."""
        missing = []
        if not (_has("pymupdf") or _has("fitz")):
            missing.append("pymupdf")
        if not _has("openpyxl"):
            missing.append("openpyxl")
        if missing:
            raise ProviderError(f"مكتبات ناقصة: {' '.join(missing)}")

        global _VOCAB
        xls = self.excel_path.get().strip()
        if Path(xls).is_file() and not _VOCAB:
            _VOCAB = learn_vocabulary(xls)
            n = sum(len(f) for f in _VOCAB.values())
            if n:
                self.log(f"📚 تُعلّمت {n} حقلاً محصوراً من المرجعية.", "info")

        prov = self._provider()
        prov.on_debug = lambda m: self.log(f"🔎 {m}", "dim")
        primary = prov.model

        def on_page(i, total):
            self.set_progress(i / total * 45)

        images = pdf_to_jpegs(pdf, progress=on_page)
        self.log(f"📄 {len(images)} صفحة — "
                 f"{sum(len(b) for b in images) / 1e6:.1f} ميغابايت", "dim")
        self.set_progress(50)

        pool = (GEMINI_MODELS if prov.name == "Gemini" else
                CLAUDE_MODELS if prov.name == "Claude" else [])
        chain = [primary] + ([m for m in pool if m != primary]
                             if self.fallback.get() else [])

        def on_retry(i, total, delay, err):
            self.log(f"⏳ ازدحام — إعادة {i}/{total - 1} بعد {delay:.0f} ثانية", "warn")

        raw = None
        for mi, mname in enumerate(chain):
            prov.model = mname
            if mi:
                self.log(f"🔄 تبديل إلى {mname}", "warn")
            try:
                raw = retry_call(lambda: prov.extract(images, build_prompt()),
                                 attempts=4, on_retry=on_retry,
                                 should_stop=lambda: self.cancel.is_set())
                break
            except Cancelled:
                raise
            except Exception as e:
                if mi == len(chain) - 1 or not is_retryable(e):
                    raise
                self.log(f"⚠️ {mname} غير متاح.", "warn")

        self.set_progress(80)
        if isinstance(raw, dict):
            data = raw
        else:
            try:
                data = parse_json_response(raw)
            except json.JSONDecodeError:
                dump = Path(pdf).with_suffix("")
                dump = dump.with_name(dump.name + "_raw_response.txt")
                try:
                    dump.write_text(str(raw), encoding="utf-8")
                    self.log(f"📄 الرد الخام: {dump.name}", "warn")
                except Exception:
                    pass
                raise

        # إصلاح أخطاء القراءة الضوئية الواضحة قبل الفحص
        fixes = normalize(data, log=lambda m: self.log(m, "warn"))
        if fixes:
            self.log(f"🔧 أُصلح {len(fixes)} حقلاً تلقائياً.", "info")
        got = enrich(data, log=lambda m: self.log(m, "dim"))
        if got:
            self.log(f"🧩 اشتُقّ {len(got)} حقلاً من معلومات متاحة.", "info")

        out = Path(pdf).with_suffix("")
        out = out.with_name(out.name + "_extracted.json")
        try:
            out.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                           encoding="utf-8")
        except Exception:
            pass
        self.set_progress(100)
        return data

    # ── الحفظ في Excel ───────────────────────────────────────────
    def start_save(self):
        if self.busy or not self.data:
            return
        xls = self.excel_path.get().strip()
        if not Path(xls).is_file():
            messagebox.showerror("خطأ", f"ملف Excel غير موجود:\n{xls}")
            return
        warns = validate(self.data)
        msg = "سيتم إنشاء نسخة احتياطية ثم إضافة البيانات.\n\nهل تريد المتابعة؟"
        if warns:
            msg = f"⚠️ هناك {len(warns)} تنبيه في البيانات:\n\n" + \
                  "\n".join("• " + w for w in warns[:6]) + \
                  ("\n…" if len(warns) > 6 else "") + "\n\n" + msg
        if not messagebox.askyesno("تأكيد الحفظ", msg):
            return
        self.busy = True
        self.b_save.config(state="disabled")
        threading.Thread(target=self._save_job, args=(xls,), daemon=True).start()

    def _save_job(self, xls: str):
        try:
            self.set_status("جارٍ الكتابة في Excel…")
            self.root.after(0, lambda: self.nb.select(self.tab_run))
            s = write_to_excel(xls, self.data, bool(self.overwrite.get()),
                               lambda m: self.log(m, "ok" if m.startswith(("✅", "💾", "♻")) else "warn"))
            n = len(s["added"]) + len(s["updated"])
            self.log(f"🎉 اكتملت العملية — {n} ورقة محدّثة.", "ok")
            self.set_status("تم الحفظ بنجاح")

            # تحديث حالة السجلّ في الطابور إن جاء منه
            if self.current_job is not None:
                self.queue.update(self.current_job,
                                  status=ST_SAVED if n else ST_SKIPPED,
                                  summary=f"{n} ورقة")
                self._file_away(self.current_job, "تم_إدخالها")
                self.current_job = None
            self.root.after(0, lambda: messagebox.showinfo(
                "تم الحفظ",
                f"✅ تمت العملية بنجاح.\n\n"
                f"أُضيفت: {len(s['added'])} ورقة\n"
                f"حُدّثت: {len(s['updated'])} ورقة\n"
                f"تُخطّيت: {len(s['skipped'])} ورقة\n\n"
                f"نسخة احتياطية:\n{Path(s['backup']).name}"))
        except PermissionError:
            self.log("❌ الملف مفتوح في Excel — أغلقه ثم أعد المحاولة.", "err")
            self.set_status("الملف مُستخدم")
            self.root.after(0, lambda: messagebox.showerror(
                "الملف مفتوح", "ملف Excel مفتوح حالياً في برنامج آخر.\nأغلقه ثم أعد المحاولة."))
        except Exception as e:
            self.log(f"❌ {e}", "err")
            self.set_status("فشل الحفظ")
            self.root.after(0, lambda: messagebox.showerror("خطأ", str(e)))
        finally:
            self.busy = False
            self.root.after(0, lambda: self.b_save.config(state="normal"))


def _has(mod: str) -> bool:
    import importlib.util
    return importlib.util.find_spec(mod) is not None


def main():
    if sys.version_info < (3, 9):
        print("يتطلب البرنامج Python 3.9 أو أحدث")
        sys.exit(1)
    root = tk.Tk()
    try:
        root.tk.call("encoding", "system", "utf-8")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
