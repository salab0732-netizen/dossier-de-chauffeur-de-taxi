#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
وحدة إعداد محاضر الدورة وجداول الإرسال
- محضر فتح الدورة
- محضر المداولة
- جداول الإرسال
- قوائم المترشحين Excel
- نقاط المترشحين
"""
import os, platform, subprocess, random
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from utils import fmt_date, fmt_raqm, norm_raqm, clean_wilaya, clean_takh, EXCEL_FILE, SHEETS, COLS_MARKAZ, COLS_MTRSH

EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "بيانات_مركز_التكوين.xlsx")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "محاضر_PDF")

def safe_filename(raqm):
    """تنظيف رقم الدورة لاستخدامه في اسم الملف — يستبدل / بـ -"""
    return str(raqm or "").replace("/", "-").replace("\\", "-").strip()

TEMPLATE_DIR   = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_MAHDAR_FATH    = os.path.join(TEMPLATE_DIR, "محضر_فتح_الدورة.docx")
TEMPLATE_MAHDAR_MUDALAWA = os.path.join(TEMPLATE_DIR, "محضر_المداولة.docx")
TEMPLATE_JADWAL_MUDALAWA = os.path.join(TEMPLATE_DIR, "جدول_ارسال_المداولة_-_طلب_اصدار_شهادة_الكفاءة_المهنية.docx")
TEMPLATE_JADWAL_FATH     = os.path.join(TEMPLATE_DIR, "فتح_الدورة_جدول_ارسال.docx")
TEMPLATE_JADWAL_FATH    = os.path.join(TEMPLATE_DIR, "فتح_الدورة_جدول_ارسال.docx")
TEMPLATE_JADWAL_MUDALWA = os.path.join(TEMPLATE_DIR, "جدول_ارسال_المداولة_-_طلب_اصدار_شهادة_الكفاءة_المهنية.docx")

# ── مواد التخصصات ────────────────────────────────────────────
MAWAD = {
    "الأشخاص": {
        "مشترك":   ["م01", "م02", "م03", "م04"],
        "نظري":    ["م01"],
        "تطبيقي":  ["م01", "م02", "م03"],
    },
    "البضائع": {
        "مشترك":   ["م01", "م02", "م03", "م04"],
        "نظري":    ["م01", "م02"],
        "تطبيقي":  ["م01", "م02"],
    },
}

def get_mawad_list(takh):
    """إرجاع قائمة أسماء المواد حسب التخصص"""
    t = str(takh or "")
    if "أشخاص" in t or "الأشخاص" in t:
        key = "الأشخاص"
    else:
        key = "البضائع"
    m = MAWAD[key]
    cols = []
    for label, items in [("مشترك", m["مشترك"]), ("نظري", m["نظري"]), ("تطبيقي", m["تطبيقي"])]:
        for item in items:
            cols.append(f"{label}/{item}")
    return cols, key


# ── التاريخ بالحروف العربية ───────────────────────────────────
DAYS_AR = {
    1:"الأول",2:"الثاني",3:"الثالث",4:"الرابع",5:"الخامس",
    6:"السادس",7:"السابع",8:"الثامن",9:"التاسع",10:"العاشر",
    11:"الحادي عشر",12:"الثاني عشر",13:"الثالث عشر",14:"الرابع عشر",
    15:"الخامس عشر",16:"السادس عشر",17:"السابع عشر",18:"الثامن عشر",
    19:"التاسع عشر",20:"العشرين",21:"الحادي والعشرين",22:"الثاني والعشرين",
    23:"الثالث والعشرين",24:"الرابع والعشرين",25:"الخامس والعشرين",
    26:"السادس والعشرين",27:"السابع والعشرين",28:"الثامن والعشرين",
    29:"التاسع والعشرين",30:"الثلاثين",31:"الحادي والثلاثين",
}
MONTHS_AR = {
    1:"جانفي",2:"فيفري",3:"مارس",4:"أفريل",5:"ماي",6:"جوان",
    7:"جويلية",8:"أوت",9:"سبتمبر",10:"أكتوبر",11:"نوفمبر",12:"ديسمبر",
}
YEARS_AR = {
    2024:"ألفين وأربعة وعشرون",
    2025:"ألفين وخمسة وعشرون",
    2026:"ألفين وستة وعشرون",
    2027:"ألفين وسبعة وعشرون",
    2028:"ألفين وثمانية وعشرون",
    2029:"ألفين وتسعة وعشرون",
    2030:"ألفين وثلاثون",
}

def date_to_arabic(d):
    """تحويل تاريخ إلى نص عربي كامل"""
    if isinstance(d, str):
        try: d = datetime.strptime(d.strip(), "%Y-%m-%d")
        except Exception: return str(d)
    day   = DAYS_AR.get(d.day, str(d.day))
    month = MONTHS_AR.get(d.month, str(d.month))
    year  = YEARS_AR.get(d.year, str(d.year))
    return f"في اليوم {day} من شهر {month} لسنة {year}"

def get_ijtimaa_date(date_str, days_before=5, no_friday=True):
    """حساب تاريخ الاجتماع قبل الدورة بأيام مع تجنب الجمعة"""
    try:
        s = str(date_str).strip().replace('/', '-')
        d = datetime.strptime(s, "%Y-%m-%d")
    except Exception:
        try:
            s = str(date_str).strip()
            # DD/MM/YYYY
            import re as _re
            m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
            if m: d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            else: return datetime.now()
        except Exception: return datetime.now()
    d = d - timedelta(days=days_before)
    if no_friday and d.weekday() == 4:  # 4 = جمعة
        d = d - timedelta(days=1)       # ننتقل للخميس
    return d

def get_ijtimaa_mudalawa(date_str, no_friday=True):
    """تاريخ اجتماع المداولة = نهاية التكوين + 1 يوم"""
    try:
        s = str(date_str).strip().replace('/', '-')
        d = datetime.strptime(s, "%Y-%m-%d")
    except Exception:
        try:
            s = str(date_str).strip()
            import re as _re
            m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
            if m: d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            else: return datetime.now()
        except Exception: return datetime.now()
    d = d + timedelta(days=1)
    if no_friday and d.weekday() == 4:  # جمعة → سبت
        d = d + timedelta(days=1)
    return d


# ── تحميل البيانات ────────────────────────────────────────────

def _fix_date(val):
    """تحويل أي صيغة تاريخ → YYYY/MM/DD دائماً"""
    if val is None: return ""
    # datetime object → استخرج المكونات مباشرة
    if hasattr(val, 'year'):
        return f"{val.year}/{str(val.month).zfill(2)}/{str(val.day).zfill(2)}"
    import re as _re
    s = str(val).strip()
    if not s or s == "None": return ""
    # YYYY/MM/DD → صحيح بالفعل
    if _re.match(r'^\d{4}/\d{2}/\d{2}$', s): return s
    # DD/MM/YYYY → YYYY/MM/DD
    m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if m: return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    # YYYY-MM-DD
    m2 = _re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m2: return f"{m2.group(1)}/{m2.group(2)}/{m2.group(3)}"
    return s

def fmt_raqm(raqm):
    """تنسيق رقم الدورة: 13/2026 → 2026/13، 1/2026 → 2026/01"""
    import re as _re
    s = str(raqm or "").strip()
    if not s: return s
    # XX/YYYY → YYYY/XX
    m = _re.match(r'^(\d+)/(\d{4})$', s)
    if m:
        return f"{m.group(2)}/{m.group(1).zfill(2)}"
    # رقم فردي → 01
    parts = s.split("/")
    fmt_parts = []
    for p in parts:
        p = p.strip()
        if p.isdigit() and len(p) == 1:
            fmt_parts.append(f"0{p}")
        else:
            fmt_parts.append(p)
    return "/".join(fmt_parts)


def _clean_ville(val):
    """يأخذ قيمة خانة بلدية المركز كما هي — بدون أي معالجة"""
    return str(val or "").strip()

def load_markaz():
    """قراءة بيانات المركز — من SQLite أولاً ثم Excel"""
    try:
        import database as _db, sqlite3 as _sq
        conn = _sq.connect(_db.DB_FILE)
        c = conn.cursor()
        c.execute("SELECT * FROM مركز LIMIT 1")
        row = c.fetchone()
        if row:
            cols = [d[0] for d in c.description]
            conn.close()
            raw = dict(zip(cols, row))
            # تحويل أسماء الأعمدة: اسم_مركز_التكوين → اسم مركز التكوين
            result = {}
            for k, v in raw.items():
                clean_key = k.replace("_", " ")
                result[clean_key] = str(v) if v is not None else ""
            if any(v for k,v in result.items() if k not in ("id","تاريخ التحديث","آخر رقم ملف")):
                return result
                return result
        conn.close()
    except Exception:
        pass
    # fallback: Excel
    if not os.path.exists(EXCEL_FILE): return {}
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["مركز_التكوين"]
    cols = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        return {k: (str(v) if v is not None else "") for k, v in zip(cols, row) if k}
    return {}

def load_dawra(raqm_dawra):
    if not os.path.exists(EXCEL_FILE): return {}
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE)
    try:
        ws = wb["الدورات"]
        raqm_str = str(raqm_dawra).strip()

        def norm(v):
            """تطبيع رقم الدورة للمقارنة — يتجاهل الأصفار البادئة"""
            s = str(v or "").strip()
            parts = s.split("/")
            normalized = []
            for p in parts:
                p = p.strip()
                try: normalized.append(str(int(p)))
                except Exception: normalized.append(p)
            return "/".join(normalized)

        raqm_norm = norm(raqm_str)

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[2] is not None:
                r_val = str(row[2]).strip()
                if r_val == raqm_str or norm(r_val) == raqm_norm:
                    return {
                        "date_ouv": _fix_date(row[0]),
                        "date_clo": _fix_date(row[1]),
                        "raqm":     r_val,
                    }
    except Exception as e:
        print(f"خطأ load_dawra: {e}")
    return {}

def load_mtrsheen(raqm_dawra):
    if not os.path.exists(EXCEL_FILE): return []
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["المترشحون"]
    # قراءة رؤوس الأعمدة من Excel مباشرة — لا تعتمد على ترتيب ثابت
    cols = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]

    # تطبيع رقم الدورة — يتجاهل الأصفار البادئة والاتجاه (YYYY/NN أو NN/YYYY)
    def norm(v):
        """تحويل أي صيغة لـ YYYY/NN للمقارنة"""
        import re as _re
        s = str(v or "").strip()
        if not s: return ""
        # تحويل NN/YYYY → YYYY/NN
        m = _re.match(r'^(\d{1,3})/(\d{4})$', s)
        if m: s = f"{m.group(2)}/{m.group(1)}"
        # إزالة الأصفار البادئة من كل جزء
        parts = s.split("/")
        normalized = []
        for p in parts:
            p = p.strip()
            try: normalized.append(str(int(p)))
            except Exception: normalized.append(p)
        return "/".join(normalized)

    raqm_norm = norm(raqm_dawra)
    results = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        d = {k: (str(v) if v is not None else "") for k, v in zip(cols, row) if k}
        dv = d.get("رقم الدورة", "").strip()
        if norm(dv) == raqm_norm:
            results.append(d)
    return results


# ── نقاط المترشحين ───────────────────────────────────────────
def gen_notes_auto(takh):
    """توليد نقاط عشوائية متنوعة بين 10 و20"""
    cols, key = get_mawad_list(takh)
    notes = {}
    for col in cols:
        # نقاط متنوعة: بعضها 10-13، بعضها 14-17، بعضها 18-20
        tier = random.choice([1, 1, 2, 2, 3])
        if tier == 1:   notes[col] = random.randint(10, 13)
        elif tier == 2: notes[col] = random.randint(14, 17)
        else:           notes[col] = random.randint(18, 20)
    moyenne = round(sum(notes.values()) / len(notes), 2)
    notes["المعدل"] = moyenne
    notes["القرار"] = "ناجح" if moyenne >= 10 else "راسب"
    return notes

def calc_moyenne(notes_dict, takh):
    """حساب المعدل من نقاط يدوية"""
    cols, _ = get_mawad_list(takh)
    vals = [float(notes_dict.get(c, 0)) for c in cols]
    if not vals: return 0
    return round(sum(vals) / len(vals), 2)


# ── حفظ النقاط في Excel ──────────────────────────────────────
def save_notes_to_excel(raqm_dawra, notes_data):
    """
    حفظ نقاط المترشحين في ورقة 'النقاط' في Excel.
    notes_data: {اسم_المتربص: {col: note, ...}, ...}
    """
    import openpyxl
    if not notes_data:
        print("⚠️ لا توجد نقاط للحفظ"); return

    wb = openpyxl.load_workbook(EXCEL_FILE)
    if "النقاط" in wb.sheetnames:
        del wb["النقاط"]
    ws = wb.create_sheet("النقاط")
    ws.sheet_view.rightToLeft = True

    # رأس الجدول
    first = next(iter(notes_data.values()))
    note_cols = [k for k in first.keys() if k not in ["المعدل","القرار","التخصص"]]
    headers = ["رقم الدورة", "الاسم واللقب", "التخصص"] + note_cols + ["المعدل", "القرار"]
    ws.append(headers)

    # تنسيق الرأس
    from openpyxl.styles import Font, PatternFill, Alignment
    hfont = Font(bold=True, color="FFFFFF", name="Arabic Typesetting", size=11)
    hfill = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = hfont; cell.fill = hfill; cell.alignment = halign
    ws.row_dimensions[1].height = 28

    # البيانات
    for i, (nom, notes) in enumerate(notes_data.items(), 2):
        row = [raqm_dawra, nom, notes.get("التخصص","")]
        for col in note_cols:
            row.append(notes.get(col, ""))
        row += [notes.get("المعدل",""), notes.get("القرار","")]
        ws.append(row)
        # تلوين حسب النجاح
        color = "E8F8E8" if notes.get("القرار") == "ناجح" else "FFF0F0"
        for cell in ws[i]:
            cell.fill = PatternFill("solid", start_color=color, end_color=color)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # عرض الأعمدة
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 15
    for i in range(4, len(headers)+1):
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(i)].width = 10

    try:
        wb.save(EXCEL_FILE)
        print(f"✅ النقاط حُفظت في Excel — {len(notes_data)} مترشح")
    except PermissionError:
        raise PermissionError("ملف Excel مفتوح — أغلقه ثم حاول مرة أخرى")


# ── قائمة المترشحين Excel ────────────────────────────────────
def export_liste_mtrsheen(raqm_dawra, output_dir=None):
    """
    تصدير قائمة المترشحين إلى Excel A4 مرفق مع محضر الفتح
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    markaz   = load_markaz()
    mtrsheen = load_mtrsheen(raqm_dawra)
    if not mtrsheen:
        print("لا يوجد مترشحون"); return None

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"قائمة_المترشحين_{safe_filename(raqm_dawra)}_{ts}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "قائمة المترشحين"
    ws.sheet_view.rightToLeft = True

    # إعداد الورقة A4
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True

    # أنماط
    bold_center = Font(name="Arabic Typesetting", bold=True, size=12)
    normal      = Font(name="Arabic Typesetting", size=11)
    header_fill = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    header_font = Font(name="Arabic Typesetting", bold=True, size=11, color="FFFFFF")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right  = Alignment(horizontal="right",  vertical="center", wrap_text=True)

    # رأس الوثيقة
    ws.merge_cells("A1:F1")
    ws["A1"] = "الجمهورية الجزائرية الديمقراطية الشعبية — وزارة الداخلية والجماعات المحلية والنقل"
    ws["A1"].font = bold_center; ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}  —  رقم الاعتماد : {markaz.get("رقم الاعتماد","")}'
    ws["A2"].font = Font(name="Arabic Typesetting", bold=True, size=11)
    ws["A2"].alignment = center

    ws.merge_cells("A3:I3")
    ws["A3"] = f"قائمة المترشحين المقبولين في الدورة رقم : {raqm_dawra}"
    ws["A3"].font = bold_center; ws["A3"].alignment = center

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 22

    # رأس الجدول
    headers = ["الرقم", "الاسم واللقب", "تاريخ ومكان الازدياد",
               "رقم التعريف الوطني", "رقم رخصة السياقة",
               "تاريخ صدور الرخصة", "فئة الرخصة", "التخصص", "رقم الملف"]
    widths  = [8, 30, 28, 22, 18, 16, 12, 15, 12]

    for col_idx, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = center; cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[4].height = 28

    # البيانات
    for i, m in enumerate(mtrsheen, 1):
        row_idx = i + 4
        takh = str(m.get("تخصص التكوين","") or "")
        data = [
            i,
            str(m.get("الاسم واللقب","") or ""),
            str(m.get("تاريخ ومكان الازدياد","") or ""),
            str(m.get("رقم التعريف الوطني","") or ""),
            str(m.get("رقم رخصة السياقة","") or ""),
            str(m.get("تاريخ صدور رخصة السياقة","") or ""),
            str(m.get("فئات الرخصة","") or ""),
            takh,
            str(m.get("رقم الملف","") or ""),
        ]
        fill_color = "EBF3FB" if i % 2 == 0 else "FFFFFF"
        for col_idx, val in enumerate(data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = normal
            cell.alignment = center if col_idx in [1,7,8,9] else right
            cell.border = border
            cell.fill = PatternFill("solid", start_color=fill_color, end_color=fill_color)
        ws.row_dimensions[row_idx].height = 22

    # إجمالي
    total_row = len(mtrsheen) + 5
    ws.merge_cells(f"A{total_row}:H{total_row}")
    ws[f"A{total_row}"] = f"إجمالي عدد المترشحين : {len(mtrsheen)} مترشح"
    ws[f"A{total_row}"].font = Font(name="Arabic Typesetting", bold=True, size=11)
    ws[f"A{total_row}"].alignment = right
    ws[f"A{total_row}"].border = border

    wb.save(out)
    print(f"✅ قائمة المترشحين → {out}")
    return out


# ── قائمة الناجحين مع النقاط Excel ──────────────────────────
def export_liste_najihin(raqm_dawra, output_dir=None):
    """
    تصدير قائمة الناجحين مع النقاط إلى Excel مرفق مع محضر المداولة
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    markaz   = load_markaz()
    mtrsheen_all = load_mtrsheen(raqm_dawra)
    # نأخذ كل المترشحين (ناجح وراسب) لمحضر المداولة
    mtrsheen = [m for m in mtrsheen_all if m.get("حالة النجاح") in ["ناجح","راسب"]]
    if not mtrsheen:
        # إذا لم توجد نتائج نأخذ الكل
        mtrsheen = mtrsheen_all
    if not mtrsheen:
        print("لا يوجد مترشحون"); return None

    # تحميل النقاط من Excel إن وجدت
    notes_map = {}
    try:
        import openpyxl as ox
        wb_n = ox.load_workbook(EXCEL_FILE)
        if "النقاط" in wb_n.sheetnames:
            ws_n = wb_n["النقاط"]
            headers_n = [c.value for c in ws_n[1]]
            for row in ws_n.iter_rows(min_row=2, values_only=True):
                d = dict(zip(headers_n, row))
                if str(d.get("رقم الدورة","")).strip() == str(raqm_dawra).strip():
                    notes_map[str(d.get("الاسم واللقب",""))] = d
    except Exception: pass

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"قائمة_الناجحين_{safe_filename(raqm_dawra)}_{ts}.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "قائمة المداولة"
    ws.sheet_view.rightToLeft = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"

    thin   = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right  = Alignment(horizontal="right",  vertical="center", wrap_text=True)
    header_fill = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    header_font = Font(name="Arabic Typesetting", bold=True, size=10, color="FFFFFF")
    normal = Font(name="Arabic Typesetting", size=10)
    bold   = Font(name="Arabic Typesetting", bold=True, size=11)

    # رأس الوثيقة
    ncols = 14
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    ws["A1"] = "الجمهورية الجزائرية الديمقراطية الشعبية — وزارة الداخلية والجماعات المحلية والنقل"
    ws["A1"].font = bold; ws["A1"].alignment = center

    ws.merge_cells(f"A2:{get_column_letter(ncols)}2")
    ws["A2"] = f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}  —  رقم الاعتماد : {markaz.get("رقم الاعتماد","")}'
    ws["A2"].font = Font(name="Arabic Typesetting", bold=True, size=11); ws["A2"].alignment = center

    ws.merge_cells(f"A3:{get_column_letter(ncols)}3")
    ws["A3"] = f"محضر المداولة — الدورة رقم : {raqm_dawra}"
    ws["A3"].font = bold; ws["A3"].alignment = center

    for r in [1,2,3]: ws.row_dimensions[r].height = 22

    # رأس الجدول
    main_headers = [
        "الرقم", "الاسم واللقب", "تاريخ ومكان الازدياد",
        "رقم التعريف الوطني", "رقم رخصة السياقة",
        "تاريخ صدور الرخصة", "فئة الرخصة",
        "التخصص", "رقم الملف",
        "م01", "م02", "م03", "المعدل", "النتيجة"
    ]
    widths = [6, 28, 26, 20, 16, 14, 10, 12, 10, 7, 7, 7, 9, 12]

    for col_idx, (h,w) in enumerate(zip(main_headers, widths), 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = center; cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[4].height = 35

    for i, m in enumerate(mtrsheen, 1):
        row_idx = i + 4
        nom  = str(m.get("الاسم واللقب","") or "")
        takh = str(m.get("تخصص التكوين","") or "")
        n    = notes_map.get(nom, {})
        moy  = n.get("المعدل", "")
        qar  = str(m.get("حالة النجاح","") or n.get("القرار","") or "")

        # تلوين حسب النتيجة
        if qar == "ناجح":
            fill_color = "E8F8E8"
        elif qar == "راسب":
            fill_color = "FFF0F0"
        else:
            fill_color = "EBF3FB" if i % 2 == 0 else "FFFFFF"

        # نأخذ أول 3 نقاط فقط للعرض المختصر
        note_keys = ["مشترك/م01","مشترك/م02","مشترك/م03"]
        data = [
            i,
            nom,
            str(m.get("تاريخ ومكان الازدياد","") or ""),
            str(m.get("رقم التعريف الوطني","") or ""),
            str(m.get("رقم رخصة السياقة","") or ""),
            str(m.get("تاريخ صدور رخصة السياقة","") or ""),
            str(m.get("فئات الرخصة","") or ""),
            takh,
            str(m.get("رقم الملف","") or ""),
        ]
        data += [n.get(k,"") for k in note_keys]
        data += [moy, qar]

        for col_idx, val in enumerate(data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = normal; cell.border = border
            cell.alignment = center if col_idx in [1,8,9,13,14] else right
            cell.fill = PatternFill("solid", start_color=fill_color, end_color=fill_color)
        ws.row_dimensions[row_idx].height = 22

    # إحصاءات
    total_row = len(mtrsheen) + 5
    najihin = sum(1 for m in mtrsheen if str(m.get("حالة النجاح","")) == "ناجح")
    rasibin = sum(1 for m in mtrsheen if str(m.get("حالة النجاح","")) == "راسب")
    ws.merge_cells(f"A{total_row}:{get_column_letter(ncols)}{total_row}")
    ws[f"A{total_row}"] = (f"الإجمالي: {len(mtrsheen)} مترشح  |  "
                           f"ناجح: {najihin}  |  راسب: {rasibin}")
    ws[f"A{total_row}"].font = Font(name="Arabic Typesetting", bold=True, size=11)
    ws[f"A{total_row}"].alignment = center
    ws[f"A{total_row}"].fill = PatternFill("solid", start_color="F0F0F0", end_color="F0F0F0")
    ws[f"A{total_row}"].border = border




def fill_jadwal_irsal_fath(raqm_dawra, output_dir=None, raqm_mahdar=None):
    """ملء جدول إرسال فتح الدورة"""
    from docx import Document as _Doc
    import shutil as _sh

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    markaz  = load_markaz()
    RLM = "‏"

    nom_markaz = str(markaz.get("اسم مركز التكوين","") or "").strip()
    wilaya     = clean_wilaya(str(markaz.get("مديرية النقل لولاية","") or "")).strip()
    ville      = _clean_ville(markaz.get("بلدية المركز",""))
    raqm_str   = fmt_raqm(str(raqm_dawra or ""))
    raqm_m     = fmt_raqm(str(raqm_mahdar or "")) if raqm_mahdar else ""

    # تاريخ جدول الإرسال = تاريخ اجتماع الافتتاح (نفس تاريخ محضر الفتح)
    dawra = load_dawra(raqm_dawra)
    date_ouv = dawra.get("date_ouv") if dawra else None
    if date_ouv:
        ijtimaa_fath = get_ijtimaa_date(str(date_ouv))
        irsal_date = ijtimaa_fath.strftime("%Y/%m/%d")
    else:
        irsal_date = datetime.now().strftime("%Y/%m/%d")

    replacements = {
        "F_MARKAZ":  RLM + nom_markaz,
        "F_WILAYA":  RLM + wilaya,
        "F_RAQM_M":  RLM + raqm_m,
        "F_RAQM_J":  RLM + raqm_m,
        "F_RAQM_D":  RLM + raqm_str,
        "F_VILLE":   RLM + ville,
        "F_DATE":    RLM + irsal_date,
    }

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"جدول_ارسال_فتح_{safe_filename(raqm_dawra)}_{ts}.docx")
    _sh.copy2(TEMPLATE_JADWAL_FATH, out)
    doc = _Doc(out)
    _replace_in_doc(doc, replacements)
    doc.save(out)
    print(f"جدول إرسال الفتح → {out}")
    return out

def fill_jadwal_irsal_mudalawa(raqm_dawra, output_dir=None, raqm_mahdar=None):
    """ملء جدول إرسال المداولة — طلب إصدار شهادة الكفاءة المهنية"""
    from docx import Document as _Doc
    import shutil as _sh

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    markaz  = load_markaz()
    mtrsheen = load_mtrsheen(raqm_dawra)
    RLM = "‏"

    # بيانات المركز
    nom_markaz = str(markaz.get("اسم مركز التكوين","") or "")
    wilaya     = clean_wilaya(str(markaz.get("مديرية النقل لولاية","") or ""))
    ville      = _clean_ville(markaz.get("بلدية المركز",""))
    raqm_str   = fmt_raqm(str(raqm_dawra or ""))
    raqm_m     = fmt_raqm(str(raqm_mahdar or "")) if raqm_mahdar else ""

    # تاريخ جدول الإرسال = تاريخ اجتماع المداولة (نفس تاريخ محضر المداولة)
    dawra_m = load_dawra(raqm_dawra)
    date_clo_m = dawra_m.get("date_clo") if dawra_m else None
    if date_clo_m:
        ijtimaa_mud = get_ijtimaa_mudalawa(str(date_clo_m))
        irsal_date_m = ijtimaa_mud.strftime("%Y/%m/%d")
    else:
        irsal_date_m = datetime.now().strftime("%Y/%m/%d")

    # إحصاء المترشحين حسب التخصص
    n_ash = sum(1 for m in mtrsheen
                if "أشخاص" in clean_takh(str(m.get("تخصص التكوين","") or ""))
                and "بضائع" not in clean_takh(str(m.get("تخصص التكوين","") or "")))
    n_bad = sum(1 for m in mtrsheen
                if "بضائع" in clean_takh(str(m.get("تخصص التكوين","") or ""))
                and "أشخاص" not in clean_takh(str(m.get("تخصص التكوين","") or "")))
    n_maz = sum(1 for m in mtrsheen
                if "أشخاص" in clean_takh(str(m.get("تخصص التكوين","") or ""))
                and "بضائع" in clean_takh(str(m.get("تخصص التكوين","") or "")))

    replacements = {
        "J_MARKAZ2":  RLM + nom_markaz,
        "J_MARKAZ":   RLM + nom_markaz,
        "J_WILAYA":   RLM + wilaya,
        "J_RAQM_M":   RLM + raqm_m,
        "J_RAQM_J":   RLM + raqm_m,
        "J_RAQM_D":   RLM + raqm_str,
        "J_VILLE":    RLM + ville,
        "J_DATE":     RLM + irsal_date_m,
        "J_NB_ASH":   RLM + str(n_ash),
        "J_NB_BAD":   RLM + str(n_bad),
        "J_NB_MAZ":   RLM + str(n_maz),
    }

    # نعمل على نسخة مؤقتة — القاعدة الذهبية
    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"جدول_ارسال_مداولة_{safe_filename(raqm_dawra)}_{ts}.docx")
    _sh.copy2(TEMPLATE_JADWAL_MUDALAWA, out)

    doc = _Doc(out)
    _replace_in_doc(doc, replacements)
    doc.save(out)
    print(f"جدول إرسال المداولة → {out}")
    return out


def open_file(path):
    if not path or not os.path.exists(path): return
    if   platform.system() == "Windows": os.startfile(path)
    elif platform.system() == "Darwin":  subprocess.Popen(["open", path])
    else:                                subprocess.Popen(["xdg-open", path])


# ── ملء نموذج Word ────────────────────────────────────────────


def _replace_in_doc(doc, replacements):
    """استبدال النص في فقرات وجداول Word — يجمع كل الـ runs أولاً"""
    def _replace_para(para, reps):
        # فحص سريع
        full = "".join(r.text or "" for r in para.runs)
        if not full: return
        if not any(key in full for key in reps): return

        # حساب مواضع الـ runs
        positions = []
        pos = 0
        for i, r in enumerate(para.runs):
            t = r.text or ""
            positions.append((pos, pos + len(t), i))
            pos += len(t)

        # استبدال في النص الكامل
        new_full = full
        for key in sorted(reps.keys(), key=len, reverse=True):
            new_full = new_full.replace(key, str(reps[key]))
        if new_full == full: return

        # هل كل علامة داخل run واحد؟
        all_single = True
        for key in reps:
            if key in full:
                idx = full.find(key)
                while idx != -1:
                    in_single = any(
                        start <= idx and (idx + len(key)) <= end
                        for (start, end, ri) in positions
                    )
                    if not in_single:
                        all_single = False
                        break
                    idx = full.find(key, idx + 1)
            if not all_single:
                break

        if all_single:
            # استبدال كل run على حدة — يحافظ على التنسيق
            for r in para.runs:
                if not r.text: continue
                t = r.text
                for key in sorted(reps.keys(), key=len, reverse=True):
                    t = t.replace(key, str(reps[key]))
                r.text = t
        else:
            # علامة مُقسَّمة — نضطر للدمج
            if para.runs:
                para.runs[0].text = new_full
                for r in para.runs[1:]: r.text = ""

    for para in doc.paragraphs:
        _replace_para(para, replacements)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _replace_para(para, replacements)


def _fmt_date(ds):
    """تحويل أي صيغة تاريخ → YYYY/MM/DD دائماً"""
    if not ds: return ""
    # datetime object → استخرج المكونات مباشرة
    if hasattr(ds, 'year'):
        return f"{ds.year}/{str(ds.month).zfill(2)}/{str(ds.day).zfill(2)}"
    s = str(ds).strip()
    if not s or s == "None": return ""
    import re as _re
    if _re.match(r'^\d{4}/\d{2}/\d{2}$', s): return s
    m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if m: return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    m2 = _re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m2: return f"{m2.group(1)}/{m2.group(2)}/{m2.group(3)}"
    return s


def fix_all_dates_in_doc(doc):
    """تحويل كل DD/MM/YYYY في الوثيقة إلى YYYY/MM/DD مع دعم الـ runs المقسّمة"""
    import re as _re

    def convert_dates(text):
        return _re.sub(
            r'\b(\d{2})/(\d{2})/(\d{4})\b',
            lambda m: f"\u200f{m.group(3)}/{m.group(2)}/{m.group(1)}",
            text
        )

    def fix_paragraph(para):
        if not para.runs: return
        # نجمع كل النص من كل الـ runs
        full = "".join(r.text or "" for r in para.runs)
        fixed = convert_dates(full)
        if fixed != full:
            # نضع النص المُصلح كاملاً في أول run ونمسح الباقي
            para.runs[0].text = fixed
            for r in para.runs[1:]:
                r.text = ""

    for para in doc.paragraphs:
        fix_paragraph(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    fix_paragraph(para)
    return doc


def fill_mahdar_fath(raqm_dawra, output_dir=None, raqm_mahdar=None):
    """ملء محضر فتح الدورة"""
    from docx import Document

    if not os.path.exists(TEMPLATE_MAHDAR_FATH):
        raise FileNotFoundError(f"النموذج غير موجود: {TEMPLATE_MAHDAR_FATH}")

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    markaz   = load_markaz()
    dawra    = load_dawra(raqm_dawra)
    mtrsheen = load_mtrsheen(raqm_dawra)

    if not dawra:
        raise ValueError(f"الدورة {raqm_dawra} غير موجودة في Excel")

    date_ouv = dawra.get("date_ouv")
    date_clo = dawra.get("date_clo")

    ijtimaa    = get_ijtimaa_date(str(date_ouv) if date_ouv else "")
    ijtimaa_ar = date_to_arabic(ijtimaa)

    total     = len(mtrsheen)
    # حساب الأولي والمتواصل من حقل "نوع التكوين"
    mutawassl = sum(1 for m in mtrsheen
                   if str(m.get("نوع التكوين","") or "").strip() == "متواصل")
    awaliyy   = total - mutawassl

    representant = str(markaz.get("ممثل مركز التكوين","") or "").strip()
    if not representant:
        representant = str(markaz.get("اسم مركز التكوين","") or "")

    import re
    # المدينة: من خانة بلدية المركز فقط
    ville = _clean_ville(markaz.get("بلدية المركز",""))

    raqm_str = str(raqm_dawra).strip()

    RLM = "‏"
    replacements = {
        "F1414": RLM + f"{ijtimaa.year}/{str(ijtimaa.month).zfill(2)}/{str(ijtimaa.day).zfill(2)}",
        "F1313": _clean_ville(markaz.get("بلدية المركز","") or ""),  # بلدية المركز
        "F1212": str(markaz.get("الأساتذة","")),
        "F1111": str(markaz.get("المسؤول البيداغوجي","")),
        "F1010": representant,
        "F999":  RLM + str(mutawassl),
        "F888":  RLM + str(awaliyy),
        "F777":  RLM + str(total),
        "F666":  RLM + _fmt_date(str(date_clo or "")),
        "F555":  RLM + _fmt_date(str(date_ouv or "")),
        "F444":  RLM + str(ijtimaa.year),
        "F333":  ijtimaa_ar.replace("في اليوم","اليوم"),
        # ⚠️ F222D يجب أن يأتي قبل F222 لأن replace("F222",...) يُخرّب F222D
        "F222D": RLM + fmt_raqm(raqm_str),                              # رقم الدورة
        "F222":  RLM + (fmt_raqm(raqm_mahdar) if raqm_mahdar else ""), # رقم المحضر
        "F111":  str(markaz.get("اسم مركز التكوين","")),
    }

    doc = Document(TEMPLATE_MAHDAR_FATH)
    _replace_in_doc(doc, replacements)
    fix_all_dates_in_doc(doc)

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"محضر_فتح_{safe_filename(raqm_dawra)}_{ts}.docx")
    doc.save(out)
    print(f"✅ محضر فتح الدورة → {out}")
    return out



def fill_mahdar_mudalawa(raqm_dawra, output_dir=None, raqm_mahdar=None):
    """ملء محضر المداولة"""
    from docx import Document

    if not os.path.exists(TEMPLATE_MAHDAR_MUDALAWA):
        raise FileNotFoundError(f"النموذج غير موجود: {TEMPLATE_MAHDAR_MUDALAWA}")

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    markaz   = load_markaz()
    dawra    = load_dawra(raqm_dawra)
    mtrsheen = load_mtrsheen(raqm_dawra)

    if not dawra:
        raise ValueError(f"الدورة {raqm_dawra} غير موجودة")

    date_clo = dawra.get("date_clo")
    ijtimaa  = get_ijtimaa_mudalawa(str(date_clo) if date_clo else "")
    ijtimaa_ar = date_to_arabic(ijtimaa)

    total    = len(mtrsheen)
    najihin  = [m for m in mtrsheen if str(m.get("حالة النجاح","") or "").strip() == "ناجح"]
    rasibin  = [m for m in mtrsheen if str(m.get("حالة النجاح","") or "").strip() == "راسب"]
    n_naj    = len(najihin)
    n_ras    = len(rasibin)
    n_ghaib  = sum(1 for m in mtrsheen if str(m.get("حالة النجاح","") or "").strip() == "غائب")
    n_wajjal = sum(1 for m in mtrsheen if str(m.get("حالة النجاح","") or "").strip() == "مؤجل")
    n_takhali= sum(1 for m in mtrsheen if str(m.get("حالة النجاح","") or "").strip() == "متخلي")
    n_muid   = sum(1 for m in mtrsheen if str(m.get("حالة النجاح","") or "").strip() == "معيد")

    from طباعة_الشهادات import _clean_takh
    n_ashkhas = sum(1 for m in najihin if "أشخاص" in _clean_takh(str(m.get("تخصص التكوين","") or "")) and "بضائع" not in _clean_takh(str(m.get("تخصص التكوين","") or "")))
    n_badaai  = sum(1 for m in najihin if "بضائع" in _clean_takh(str(m.get("تخصص التكوين","") or "")) and "أشخاص" not in _clean_takh(str(m.get("تخصص التكوين","") or "")))
    n_mazdouj = sum(1 for m in najihin if "أشخاص" in _clean_takh(str(m.get("تخصص التكوين","") or "")) and "بضائع" in _clean_takh(str(m.get("تخصص التكوين","") or "")))

    representant = str(markaz.get("ممثل مركز التكوين","") or "").strip()
    if not representant:
        representant = str(markaz.get("اسم مركز التكوين","") or "")

    import re
    # المدينة: من خانة بلدية المركز فقط
    ville = _clean_ville(markaz.get("بلدية المركز",""))
    raqm_str = str(raqm_dawra).strip()

    # الأطول أولاً لتجنب التداخل
    RLM = "‏"
    replacements = {
        "M1414": RLM + f"{ijtimaa.year}/{str(ijtimaa.month).zfill(2)}/{str(ijtimaa.day).zfill(2)}",
        "M1313": _clean_ville(markaz.get("بلدية المركز","") or ""),  # بلدية المركز
        "M1212": RLM + str(n_mazdouj),
        "M1111": RLM + str(n_ashkhas),
        "M1010": RLM + str(n_badaai),
        "M999":  str(markaz.get("الأساتذة","")),
        "M888":  str(markaz.get("المسؤول البيداغوجي","")),
        "M777":  representant,
        "M666":  RLM + str(n_naj),
        "M555":  RLM + str(n_ras),
        "M_GAB": RLM + str(n_ghaib),
        "M_WAJ": RLM + str(n_wajjal),
        "M_TKH": RLM + str(n_takhali),
        "M_MID": RLM + str(n_muid),
        "M444":  RLM + str(total),
        "M333":  ijtimaa_ar.replace("في اليوم","اليوم"),
        # ⚠️ M222D يجب أن يأتي قبل M222 لأن replace("M222",...) يُخرّب M222D
        "M222D": RLM + fmt_raqm(raqm_str),                              # رقم الدورة
        "M222":  RLM + (fmt_raqm(raqm_mahdar) if raqm_mahdar else ""), # رقم المحضر
        "M111":  str(markaz.get("اسم مركز التكوين","")),
    }

    doc = Document(TEMPLATE_MAHDAR_MUDALAWA)
    _replace_in_doc(doc, replacements)
    fix_all_dates_in_doc(doc)

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(out_dir, f"محضر_مداولة_{safe_filename(raqm_dawra)}_{ts}.docx")
    doc.save(out)
    print(f"✅ محضر المداولة → {out}")
    return out


def migrate_mtrsheen():
    """
    ترحيل بيانات المترشحين من الهيكل القديم (16 عمود) للهيكل الجديد (25 عمود)
    يُشغَّل مرة واحدة فقط
    """
    import openpyxl
    if not os.path.exists(EXCEL_FILE):
        print("❌ ملف Excel غير موجود"); return 0

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["المترشحون"]

    # رأس الجدول الجديد
    new_headers = [
        "تاريخ التسجيل","الاسم واللقب","اللقب بالفرنسية","الاسم بالفرنسية",
        "تاريخ ومكان الازدياد","العنوان","رقم التعريف الوطني",
        "رقم رخصة السياقة","تاريخ صدور رخصة السياقة","تاريخ انتهاء رخصة السياقة",
        "مكان إصدار الرخصة","فئات الرخصة","زمرة الدم","الجنس",
        "رقم هاتف المتربص","البريد الإلكتروني","تخصص التكوين","رقم الدورة","نوع التكوين",
        "حالة النجاح","رقم الشهادة","تاريخ الشهادة",
        "رقم الشهادة 2","تاريخ الشهادة 2",
        "مسار صورة الرخصة","مسار صورة الشهادة"
    ]

    # قراءة كل الصفوف
    all_rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not all_rows:
        print("لا توجد بيانات للترحيل"); return 0

    # تحديث رأس الجدول
    for i, h in enumerate(new_headers, 1):
        ws.cell(row=1, column=i, value=h)

    # الهيكل القديم:
    # 0=تاريخ, 1=اسم, 2=تاريخ_ازدياد, 3=عنوان, 4=NIN, 5=رخصة, 6=تاريخ_رخصة
    # 7=تخصص, 8=هاتف, 9=بريد, 10=دورة, 11=حالة, 12=شهادة, 13=تاريخ_شهادة
    # 14=شهادة2, 15=تاريخ_شهادة2

    migrated = 0
    for row_idx, row in enumerate(all_rows, 2):
        if not any(row): continue
        n = len(row)

        # تحديد إذا كان الهيكل قديماً
        # الهيكل القديم: العمود 11 (index 10) = رقم الدورة
        # الهيكل الجديد: العمود 18 (index 17) = رقم الدورة
        is_old = n <= 16 or (n > 10 and (row[17] if n > 17 else "") == "")

        if is_old and n >= 11:
            # خريطة الترحيل: موضع_قديم → موضع_جديد
            new_row = [""] * 25
            new_row[0]  = row[0] if n > 0 else ""   # تاريخ التسجيل
            new_row[1]  = row[1] if n > 1 else ""   # الاسم واللقب
            new_row[2]  = ""                          # اللقب بالفرنسية (جديد)
            new_row[3]  = ""                          # الاسم بالفرنسية (جديد)
            new_row[4]  = row[2] if n > 2 else ""   # تاريخ ومكان الازدياد
            new_row[5]  = row[3] if n > 3 else ""   # العنوان
            new_row[6]  = row[4] if n > 4 else ""   # رقم التعريف الوطني
            new_row[7]  = row[5] if n > 5 else ""   # رقم رخصة السياقة
            new_row[8]  = row[6] if n > 6 else ""   # تاريخ صدور رخصة السياقة
            new_row[9]  = ""                          # تاريخ انتهاء الرخصة (جديد)
            new_row[10] = ""                          # مكان إصدار الرخصة (جديد)
            new_row[11] = ""                          # فئات الرخصة (جديد)
            new_row[12] = ""                          # زمرة الدم (جديد)
            new_row[13] = ""                          # الجنس (جديد)
            new_row[14] = row[8] if n > 8 else ""   # رقم هاتف المتربص
            new_row[15] = row[9] if n > 9 else ""   # البريد الإلكتروني
            new_row[16] = row[7] if n > 7 else ""   # تخصص التكوين
            new_row[17] = row[10] if n > 10 else "" # رقم الدورة ← المهم!
            new_row[18] = row[11] if n > 11 else "" # حالة النجاح
            new_row[19] = row[12] if n > 12 else "" # رقم الشهادة
            new_row[20] = row[13] if n > 13 else "" # تاريخ الشهادة
            new_row[21] = row[14] if n > 14 else "" # رقم الشهادة 2
            new_row[22] = row[15] if n > 15 else "" # تاريخ الشهادة 2
            new_row[23] = ""                          # مسار صورة الرخصة
            new_row[24] = ""                          # مسار صورة الشهادة

            for i, val in enumerate(new_row, 1):
                ws.cell(row=row_idx, column=i, value=val)
            migrated += 1
            print(f"  ✅ ترحيل: {new_row[1]} → دورة {new_row[17]}")

    if migrated > 0:
        wb.save(EXCEL_FILE)
        print(f"✅ تم ترحيل {migrated} مترشح بنجاح")
    else:
        print("✅ البيانات محدَّثة مسبقاً — لا حاجة للترحيل")
    return migrated