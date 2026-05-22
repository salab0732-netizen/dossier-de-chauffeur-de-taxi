#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
utils.py — دوال وثوابت مشتركة بين ملفات المشروع
"""
import os, re
from datetime import datetime
from dateutil.relativedelta import relativedelta

# ── مسارات ───────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "بيانات_مركز_التكوين.xlsx")

# ── أوراق Excel ───────────────────────────────────────────────
SHEETS = {"مركز": "مركز_التكوين", "دورات": "الدورات",
          "مترشحون": "المترشحون", "نقاط": "النقاط"}

# ── أعمدة الجداول ────────────────────────────────────────────
COLS_MARKAZ = ["اسم مركز التكوين","رقم الاعتماد","تاريخ الاعتماد","عدد الأقسام",
               "مديرية النقل لولاية","عنوان مركز التكوين","بلدية المركز",
               "ممثل مركز التكوين","المسؤول البيداغوجي","الأساتذة",
               "رقم الهاتف","البريد الإلكتروني","آخر رقم ملف"]

COLS_DAWRA  = ["تاريخ افتتاح الدورة","تاريخ اختتام الدورة","رقم الدورة",
               "العدد الإجمالي للمترشحين","عدد المترشحين لنقل الأشخاص",
               "عدد المترشحين لنقل البضائع","عدد المترشحين لنقل المواد الخطرة",
               "عدد المتابطين","عدد الأقسام","رقم المحضر","حالة الدورة"]

COLS_MTRSH  = ["تاريخ التسجيل","الاسم واللقب","اللقب بالفرنسية","الاسم بالفرنسية",
               "تاريخ ومكان الازدياد","العنوان","رقم التعريف الوطني",
               "رقم رخصة السياقة","تاريخ صدور رخصة السياقة","تاريخ انتهاء رخصة السياقة",
               "مكان إصدار الرخصة","فئات الرخصة","زمرة الدم","الجنس",
               "رقم هاتف المتربص","البريد الإلكتروني","تخصص التكوين","رقم الدورة",
               "نوع التكوين","حالة النجاح","رقم الشهادة","تاريخ الشهادة",
               "رقم الشهادة 2","تاريخ الشهادة 2","مسار صورة الرخصة",
               "مسار صورة الشهادة","رقم الملف"]

# ── تنسيق التاريخ ────────────────────────────────────────────
def _parse_date(ds):
    """تحويل أي صيغة تاريخ إلى كائن datetime"""
    s = str(ds).strip()
    if not s or s.lower() == "none": return None
    if hasattr(ds, 'strftime'): return ds
    if re.match(r'^\d{4}/\d{2}/\d{2}$', s):
        try: return datetime.strptime(s, "%Y/%m/%d")
        except Exception: pass
    try: return datetime.strptime(s.split(" ")[0], "%Y-%m-%d")
    except Exception: pass
    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if m:
        try: return datetime.strptime(s, "%d/%m/%Y")
        except Exception: pass
    return None

def fmt_date(ds):
    """تحويل أي صيغة تاريخ → YYYY/MM/DD"""
    if not ds: return ""
    d = _parse_date(ds)
    if d: return d.strftime("%Y/%m/%d")
    return str(ds).strip()

def calc_exp(ds, years=5):
    """حساب تاريخ انتهاء الصلاحية (افتراضياً 5 سنوات)"""
    if not ds: return ""
    d = _parse_date(ds)
    if d: return (d + relativedelta(years=years)).strftime("%Y/%m/%d")
    return str(ds)

# ── تنسيق رقم الدورة ─────────────────────────────────────────
def fmt_raqm(raqm):
    """تطبيع رقم الدورة → YYYY/N (بدون أصفار بادئة)"""
    if not raqm: return ""
    s = str(raqm).strip()
    m = re.match(r'^(\d{1,3})/(\d{4})$', s)
    if m: return f"{m.group(2)}/{m.group(1)}"
    return s

def norm_raqm(v):
    """تطبيع رقم الدورة للمقارنة — يتجاهل الأصفار البادئة"""
    s = str(v or "").strip()
    if not s: return ""
    m = re.match(r'^(\d{1,3})/(\d{4})$', s)
    if m: s = f"{m.group(2)}/{m.group(1)}"
    parts = s.split("/")
    return "/".join(str(int(p)) if p.strip().isdigit() else p.strip() for p in parts)

# ── تنظيف النصوص ─────────────────────────────────────────────
def clean(val):
    """إزالة علامات الاتجاه غير المرئية"""
    return str(val).replace("\u200f","").replace("\u200e","").strip()

def clean_wilaya(w):
    """إزالة الرقم والشرطة من اسم الولاية — '30-ورقلة' → 'ورقلة'"""
    return re.sub(r'^\d+-', '', str(w or '')).strip()

def clean_takh(takh):
    """تطبيع اسم التخصص للعرض والمقارنة"""
    return (str(takh)
            .replace("نقل الأشخاص و البضائع", "أشخاص/بضائع")
            .replace("نقل الأشخاص",            "أشخاص")
            .replace("نقل البضائع",             "بضائع")
            .replace("نقل المواد الخطرة",       "مواد خطرة")
            .replace("الأشخاص و البضائع",       "أشخاص/بضائع")
            .replace("الأشخاص", "أشخاص")
            .replace("البضائع", "بضائع")
            .replace("المواد الخطرة", "مواد خطرة"))
