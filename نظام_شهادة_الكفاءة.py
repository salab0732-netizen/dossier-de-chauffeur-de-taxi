#!/usr/bin/env python3
# -*- coding: utf-8 -*-
VERSION = "v18 — 2026-05-21"
import sys, os, traceback

# تسجيل الأخطاء في ملف log
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_log.txt")
def log_error(e):
    with open(log_file, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, webbrowser, subprocess, sys, urllib.parse
from datetime import datetime
from dateutil.relativedelta import relativedelta

# ── تهيئة قاعدة البيانات SQLite ──────────────────────────────
try:
    import database as db
    db.init_db()
    _USE_DB = True
    print("✅ SQLite جاهز")
except Exception as _db_err:
    print(f"⚠️ SQLite: {_db_err}")
    _USE_DB = False

BG="#1e2130"; BG2="#262b3d"; BG3="#2e3450"; ACCENT="#4a9eff"
ACCENT2="#f5c842"; TEXT="#e8eaf0"; TEXT2="#9ba3c0"
SUCCESS="#2ecc71"; DANGER="#e74c3c"; WARNING="#f39c12"; WHITE="#ffffff"
import platform as _platform
FONT_EMOJI = "Segoe UI Emoji" if _platform.system()=="Windows" else "Arial"
def _efont(size=10, bold=False):
    w = "bold" if bold else "normal"
    return (FONT_EMOJI, size, w)

# مسار قاعدة البيانات بجانب ملف البرنامج دائماً
EXCEL_FILE=os.path.join(os.path.dirname(os.path.abspath(__file__)),"بيانات_مركز_التكوين.xlsx")
SHEETS={"مركز":"مركز_التكوين","دورات":"الدورات","مترشحون":"المترشحون"}
COLS_MARKAZ=["اسم مركز التكوين","رقم الاعتماد","تاريخ الاعتماد","عدد الأقسام","مديرية النقل لولاية","عنوان مركز التكوين","بلدية المركز","ممثل مركز التكوين","المسؤول البيداغوجي","الأساتذة","رقم الهاتف","البريد الإلكتروني","آخر رقم ملف"]
COLS_DAWRA=["تاريخ افتتاح الدورة","تاريخ اختتام الدورة","رقم الدورة","العدد الإجمالي للمترشحين","عدد المترشحين لنقل الأشخاص","عدد المترشحين لنقل البضائع","عدد المترشحين لنقل المواد الخطرة","عدد المتابطين","عدد الأقسام","رقم المحضر","حالة الدورة"]
COLS_MTRSH=[
    "تاريخ التسجيل",          # 0
    "الاسم واللقب",            # 1
    "اللقب بالفرنسية",         # 2
    "الاسم بالفرنسية",         # 3
    "تاريخ ومكان الازدياد",    # 4
    "العنوان",                 # 5
    "رقم التعريف الوطني",      # 6
    "رقم رخصة السياقة",        # 7
    "تاريخ صدور رخصة السياقة", # 8
    "تاريخ انتهاء رخصة السياقة",# 9
    "مكان إصدار الرخصة",       # 10
    "فئات الرخصة",             # 11
    "زمرة الدم",               # 12
    "الجنس",                   # 13
    "رقم هاتف المتربص",        # 14
    "البريد الإلكتروني",       # 15
    "تخصص التكوين",            # 16
    "رقم الدورة",              # 17
    "نوع التكوين",             # 18  أولي / متواصل
    "حالة النجاح",             # 19
    "رقم الشهادة",             # 20
    "تاريخ الشهادة",           # 21
    "رقم الشهادة 2",           # 22
    "تاريخ الشهادة 2",         # 23
    "مسار صورة الرخصة",        # 24
    "مسار صورة الشهادة",       # 25
    "رقم الملف",               # 26 ← رقم الملف الوحيد لكل وثائق المترشح
]
TAKHASSAS=["الأشخاص","البضائع","المواد الخطرة","الأشخاص و البضائع"]
VALIDITY=5

WILAYAS = [
    "01-أدرار","02-الشلف","03-الأغواط","04-أم البواقي","05-باتنة",
    "06-بجاية","07-بسكرة","08-بشار","09-البليدة","10-البويرة",
    "11-تمنراست","12-تبسة","13-تلمسان","14-تيارت","15-تيزي وزو",
    "16-الجزائر العاصمة","17-الجلفة","18-جيجل","19-سطيف","20-سعيدة",
    "21-سكيكدة","22-سيدي بلعباس","23-عنابة","24-قالمة","25-قسنطينة",
    "26-المدية","27-مستغانم","28-المسيلة","29-معسكر","30-ورقلة",
    "31-وهران","32-البيض","33-إليزي","34-برج بوعريريج","35-بومرداس",
    "36-الطارف","37-تندوف","38-تيسمسيلت","39-الوادي","40-خنشلة",
    "41-سوق أهراس","42-تيبازة","43-ميلة","44-عين الدفلى","45-النعامة",
    "46-عين تموشنت","47-غرداية","48-غليزان",
    "49-تيميمون","50-برج باجي مختار","51-أولاد جلال","52-بني عباس",
    "53-عين صالح","54-عين قزام","55-تقرت","56-جانت","57-المغير","58-المنيعة",
    "59-آفلو","60-بريكة","61-القنطرة","62-بير العاتر",
    "63-العريشة","64-قصر الشلالة","65-عين وسارة","66-مسعد",
    "67-قصر البخاري","68-بوسعادة","69-الأبيض سيدي الشيخ",
]


def _hstyle():
    return (Font(name="Arial",bold=True,color="FFFFFF",size=10),
            PatternFill("solid",start_color="1F3864"),
            Alignment(horizontal="center",vertical="center",wrap_text=True),
            Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin")))

def init_excel():
    if os.path.exists(EXCEL_FILE):
        _migrate(); return
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    hf,hfill,halign,border=_hstyle()
    for sname,cols in [(SHEETS["مركز"],COLS_MARKAZ),(SHEETS["دورات"],COLS_DAWRA),(SHEETS["مترشحون"],COLS_MTRSH)]:
        ws=wb.create_sheet(sname); ws.sheet_view.rightToLeft=True; ws.row_dimensions[1].height=30
        for i,col in enumerate(cols,1):
            c=ws.cell(row=1,column=i,value=col); c.font=hf; c.fill=hfill; c.alignment=halign; c.border=border
            ws.column_dimensions[get_column_letter(i)].width=22
        ws.freeze_panes="A2"
    wb.save(EXCEL_FILE)

def _migrate():
    try:
        wb=openpyxl.load_workbook(EXCEL_FILE); changed=False
        hf,hfill,halign,border=_hstyle()
        ws_m=wb[SHEETS["مركز"]]; ex=[ws_m.cell(1,i).value for i in range(1,ws_m.max_column+1)]
        for col_name in ["رقم الهاتف","البريد الإلكتروني"]:
            if col_name not in ex:
                idx=ws_m.max_column+1; c=ws_m.cell(row=1,column=idx,value=col_name)
                c.font=hf; c.fill=hfill; c.alignment=halign; ws_m.column_dimensions[get_column_letter(idx)].width=22; changed=True
        # إضافة بلدية المركز في الموضع الصحيح (بعد عنوان المركز = عمود 6)
        if "بلدية المركز" not in ex:
            ws_m.insert_cols(7)
            c = ws_m.cell(row=1, column=7, value="بلدية المركز")
            c.font=hf; c.fill=hfill; c.alignment=halign
            ws_m.column_dimensions[get_column_letter(7)].width=22
            changed=True
        ws_t=wb[SHEETS["مترشحون"]]; ex_t=[ws_t.cell(1,i).value for i in range(1,ws_t.max_column+1)]
        for col_name in ["رقم الشهادة 2","تاريخ الشهادة 2","نوع التكوين"]:
            if col_name not in ex_t:
                # نوع التكوين يجب إدراجه في الموضع الصحيح (بعد رقم الدورة)
                if col_name == "نوع التكوين" and "نوع التكوين" not in ex_t:
                    # إدراج العمود في الموضع 19 (بعد رقم الدورة عمود 18)
                    ws_t.insert_cols(19)
                    c = ws_t.cell(row=1, column=19, value="نوع التكوين")
                    c.font=hf; c.fill=hfill; c.alignment=halign
                    ws_t.column_dimensions[get_column_letter(19)].width=18
                    # ملء القيم الموجودة بـ "أولي" افتراضياً
                    for row_i in range(2, ws_t.max_row+1):
                        if ws_t.cell(row=row_i, column=1).value:  # صف غير فارغ
                            ws_t.cell(row=row_i, column=19).value = "أولي"
                else:
                    idx=ws_t.max_column+1; c=ws_t.cell(row=1,column=idx,value=col_name)
                    c.font=hf; c.fill=hfill; c.alignment=halign; ws_t.column_dimensions[get_column_letter(idx)].width=22
                changed=True
        # إضافة حالة الدورة في ورقة الدورات
        ws_d = wb[SHEETS["دورات"]]
        ex_d = [ws_d.cell(1,i).value for i in range(1, ws_d.max_column+1)]
        if "حالة الدورة" not in ex_d:
            idx_d = ws_d.max_column + 1
            c = ws_d.cell(row=1, column=idx_d, value="حالة الدورة")
            c.font=hf; c.fill=hfill; c.alignment=halign
            ws_d.column_dimensions[get_column_letter(idx_d)].width = 18
            # ملء الدورات الموجودة بـ "مفتوحة" افتراضياً
            for row_i in range(2, ws_d.max_row+1):
                if ws_d.cell(row_i, 3).value:  # رقم الدورة
                    ws_d.cell(row_i, idx_d).value = "مفتوحة"
            changed = True
        if changed: wb.save(EXCEL_FILE)
    except Exception: pass

def load_sheet(k):
    if not os.path.exists(EXCEL_FILE): init_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[SHEETS[k]]
    result = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(c is not None for c in row): continue
        new_row = []
        for val in row:
            # تحويل datetime مباشرة
            if hasattr(val, 'strftime'):
                val = val.strftime("%Y/%m/%d")
            elif val is not None:
                s = str(val).strip()
                # تحويل DD/MM/YYYY → YYYY/MM/DD
                import re as _re
                m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
                if m: val = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
                # تحويل YYYY-MM-DD → YYYY/MM/DD
                m2 = _re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
                if m2: val = f"{m2.group(1)}/{m2.group(2)}/{m2.group(3)}"
            new_row.append(val)
        result.append(new_row)
    return result

def _update_dawra_count_excel(raqm_dawra):
    """تحديث عدد المترشحين في ورقة الدورات — يُستدعى تلقائياً بعد كل حفظ مترشح"""
    if not os.path.exists(EXCEL_FILE): return
    import re as _re
    def _norm(v):
        s = str(v or "").strip()
        if not s: return ""
        m = _re.match(r'^(\d{1,3})/(\d{4})$', s)
        if m: s = f"{m.group(2)}/{m.group(1)}"
        parts = s.split("/")
        return "/".join(str(int(p)) if p.strip().isdigit() else p.strip() for p in parts)

    raqm_norm = _norm(raqm_dawra)
    wb = openpyxl.load_workbook(EXCEL_FILE)

    # عدّ المترشحين في ورقة المترشحون
    ws_m = wb[SHEETS["مترشحون"]]
    hm = [ws_m.cell(1, i).value for i in range(1, ws_m.max_column + 1)]
    count = 0
    try:
        idx_d = hm.index("رقم الدورة")
        for row in ws_m.iter_rows(min_row=2, values_only=True):
            if not any(row): continue
            dv = str(row[idx_d] or "").strip()
            if _norm(dv) == raqm_norm:
                count += 1
    except (ValueError, IndexError):
        pass

    # تحديث العمود 4 (العدد الإجمالي) في ورقة الدورات
    ws_d = wb[SHEETS["دورات"]]
    hd = [ws_d.cell(1, i).value for i in range(1, ws_d.max_column + 1)]
    try:
        idx_r = hd.index("رقم الدورة")        # عمود رقم الدورة
        idx_c = hd.index("العدد الإجمالي للمترشحين")  # عمود العدد
        for row_i in range(2, ws_d.max_row + 1):
            dv = str(ws_d.cell(row_i, idx_r + 1).value or "").strip()
            if _norm(dv) == raqm_norm:
                ws_d.cell(row_i, idx_c + 1).value = count
                break
    except (ValueError, IndexError):
        pass

    wb.save(EXCEL_FILE)


def save_row(k, data):
    """حفظ في SQLite + Excel معاً"""
    # ── Excel ────────────────────────────────────────────────
    if not os.path.exists(EXCEL_FILE): init_excel()
    wb=openpyxl.load_workbook(EXCEL_FILE); ws=wb[SHEETS[k]]
    rf=PatternFill("solid",start_color="EBF5FB"); af=PatternFill("solid",start_color="FDFEFE")
    border=Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
    nr=ws.max_row+1; fill=rf if nr%2==0 else af
    # الأعمدة التي يجب حفظها كنص (للحفاظ على الأصفار البادئة)
    text_cols = {27}  # رقم الملف (عمود 27) — يُحفظ كنص للحفاظ على الأصفار البادئة
    for i,val in enumerate(data,1):
        cell=ws.cell(row=nr,column=i,value=val); cell.fill=fill; cell.border=border
        cell.alignment=Alignment(horizontal="right",vertical="center")
        # إجبار الخلية على قراءة القيمة كنص إذا كانت تحتوي أصفاراً بادئة
        if i in text_cols and val and str(val).startswith("0"):
            cell.number_format = "@"
            cell.value = str(val)
    wb.save(EXCEL_FILE)

    # ── تحديث عدد المترشحين في ورقة الدورات ─────────────────
    if k == "مترشحون":
        try:
            raqm_d = str(data[17] or "").strip() if len(data) > 17 else ""
            if raqm_d:
                _update_dawra_count_excel(raqm_d)
        except Exception as e:
            print(f"update_count error: {e}")

    # ── SQLite ─────────────────────────────────────────────
    if _USE_DB:
        try:
            if k == "مترشحون":
                cols = COLS_MTRSH
                d = dict(zip(cols, list(data)+[""]*(len(cols)-len(data))))
                db.save_mtrsh(d)
            elif k == "دورات":
                db.save_dawra(str(data[2] or ""), str(data[0] or ""), str(data[1] or ""), str(data[9] or "") if len(data) > 9 else "")
            elif k == "مركز":
                cols = COLS_MARKAZ
                d = dict(zip(cols, list(data)+[""]*(len(cols)-len(data))))
                db.save_markaz(d)
        except Exception as e:
            print(f"SQLite save_row error: {e}")

def get_markaz():
    rows=load_sheet("مركز"); return rows[0] if rows else [""]*len(COLS_MARKAZ)

def _get_dawra_maftuha():
    """إرجاع رقم الدورة المفتوحة للتسجيل (حالة = مفتوحة)"""
    try:
        rows = load_sheet("دورات")
        for r in rows:
            if not any(r): continue
            n = len(r)
            hala = str(r[10] or "مفتوحة").strip() if n > 10 else "مفتوحة"
            if hala == "مفتوحة":
                return fmt_raqm(str(r[2] or "")) if n > 2 else ""
    except Exception as e:
        print(f"_get_dawra_maftuha: {e}")
    return ""


def _set_dawra_hala(raqm_dawra, hala):
    """تحديث حالة الدورة في Excel: مفتوحة / جارية / منتهية"""
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb[SHEETS["دورات"]]
        # تحديد عمود حالة الدورة من الرأس
        headers = [ws.cell(1,i).value for i in range(1, ws.max_column+1)]
        hala_col = headers.index("حالة الدورة") + 1 if "حالة الدورة" in headers else None
        if not hala_col: return
        for row in ws.iter_rows(min_row=2):
            if fmt_raqm(str(row[2].value or "")) == fmt_raqm(str(raqm_dawra)):
                row[hala_col-1].value = hala
                break
        wb.save(EXCEL_FILE)
    except Exception as e:
        print(f"_set_dawra_hala: {e}")


def get_markaz_dict():
    """قراءة بيانات المركز — يقرأ رؤوس الأعمدة من Excel مباشرة لتجنب أي إزاحة"""
    if not os.path.exists(EXCEL_FILE): return {}
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[SHEETS["مركز"]]
    cols = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        return {k: (str(v) if v is not None else "") for k, v in zip(cols, row) if k}
    return {}

def next_file_num():
    """توليد رقم الملف التالي بصيغة YYYY/NNNN مثل 2026/0024"""
    year = str(datetime.now().year)
    if _USE_DB:
        try:
            num = db.next_file_num_db()  # يُرجع "0024"
            # إذا كان يحتوي سنة بالفعل نُرجعه كما هو
            if "/" in str(num): return str(num)
            return f"{year}/{str(num).zfill(4)}"
        except Exception as e:
            print(f"next_file_num DB error: {e}")
    # احتياطي من Excel
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb[SHEETS["مركز"]]
        last = 0
        for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
            if len(row) > 11 and row[11]:
                try: last = int(str(row[11]).strip().split("/")[-1])
                except Exception: last = 0
        new_num = last + 1
        for row in ws.iter_rows(min_row=2, max_row=2):
            if len(row) > 11:
                row[11].value = new_num
                break
        wb.save(EXCEL_FILE)
        return f"{year}/{str(new_num).zfill(4)}"
    except Exception as e:
        print(f"next_file_num error: {e}")
        return next_num(12)

def next_num(col=12):
    rows=load_sheet("مترشحون"); nums=[]
    for r in rows:
        if len(r)>col and r[col]:
            try: nums.append(int(str(r[col]).split("/")[-1]))
            except Exception: pass
    return str(datetime.now().year)+"/"+str(max(nums,default=0)+1).zfill(4)

def calc_exp(ds):
    if not ds: return ""
    try: return (datetime.strptime(str(ds),"%Y-%m-%d")+relativedelta(years=VALIDITY)).strftime("%Y/%m/%d")
    except Exception: return str(ds)

def fmt_date(ds):
    """تحويل أي صيغة تاريخ → YYYY/MM/DD"""
    if not ds: return ""
    import re as _re
    s = str(ds).strip()
    if not s or s == "None": return ""
    # YYYY/MM/DD — صحيح بالفعل
    if _re.match(r'^\d{4}/\d{2}/\d{2}$', s): return s
    # DD/MM/YYYY → YYYY/MM/DD
    m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if m: return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    # YYYY/MM/DD
    try: return datetime.strptime(s.split(" ")[0], "%Y-%m-%d").strftime("%Y/%m/%d")
    except Exception: pass
    # datetime object
    if hasattr(ds,'strftime'): return ds.strftime("%Y/%m/%d")
    return s

def sbtn(parent, text, cmd, bg=None, **kw):
    """زر مع دعم RTL صحيح: يفصل الإيموجي عن النص العربي"""
    kw.setdefault("pady", 7)
    kw.setdefault("padx", 12)
    bg = bg or ACCENT
    import re as _re
    # استخراج الإيموجي (أول حرف إذا كان إيموجي) والنص العربي
    text = str(text).replace("\u200f", "").strip()
    m = _re.match(r'^([\U00010000-\U0010ffff\u2000-\u2bff\u2600-\u27ff]+\s*)(.*)', text)
    if m:
        ico  = m.group(1).strip()
        lbl  = m.group(2).strip()
    else:
        ico, lbl = "", text

    if ico and lbl:
        # إطار يجمع الإيموجي والنص
        frame = tk.Frame(parent, bg=bg, cursor="hand2",
                         pady=kw.get("pady",7), padx=kw.get("padx",4))
        tk.Label(frame, text=lbl,  bg=bg, fg=WHITE,
                 font=("Arial", 10, "bold")).pack(side="right", padx=(0,4))
        tk.Label(frame, text=ico,  bg=bg, fg=WHITE,
                 font=_efont(11)).pack(side="right")
        frame.bind("<Button-1>", lambda e: cmd())
        for w in frame.winfo_children():
            w.bind("<Button-1>", lambda e: cmd())
        return frame
    else:
        return tk.Button(parent, text=text, command=cmd,
                         font=("Arial", 10, "bold"),
                         bg=bg, fg=WHITE, relief="flat",
                         cursor="hand2", **kw)

def mk_tree(parent, cols, h=12):
    s = ttk.Style()
    s.configure("Treeview", background=BG2, foreground=TEXT, rowheight=26,
                fieldbackground=BG2, font=("Arial",9))
    s.configure("Treeview.Heading", background=BG3, foreground=ACCENT2,
                font=("Arial",9,"bold"))
    t  = ttk.Treeview(parent, columns=cols, show="headings", height=h)
    sb = ttk.Scrollbar(parent, orient="vertical",   command=t.yview)
    sh = ttk.Scrollbar(parent, orient="horizontal", command=t.xview)
    t.configure(yscrollcommand=sb.set, xscrollcommand=sh.set)

    # ── تفعيل سحب الأعمدة ────────────────────────────────────
    def _start_resize(event):
        region = t.identify_region(event.x, event.y)
        if region == "separator":
            t._resize_col  = t.identify_column(event.x)
            t._resize_x    = event.x
            t._resize_w    = t.column(t._resize_col, "width")
            t.config(cursor="sb_h_double_arrow")

    def _do_resize(event):
        if hasattr(t, "_resize_col") and t._resize_col:
            dx  = event.x - t._resize_x
            new = max(40, t._resize_w + dx)
            t.column(t._resize_col, width=new)

    def _stop_resize(event):
        if hasattr(t, "_resize_col"):
            t._resize_col = None
            t.config(cursor="")

    t.bind("<ButtonPress-1>",   _start_resize)
    t.bind("<B1-Motion>",       _do_resize)
    t.bind("<ButtonRelease-1>", _stop_resize)

    return t, sb


def rtl(val):
    """عرض الأرقام والتواريخ من اليمين لليسار بشكل صحيح"""
    import re
    s = str(val).strip() if val else ""
    if not s: return s
    # إذا كان رقماً أو تاريخاً يحتوي / أو - نعكسه
    if re.match(r'^[\d/\-. ]+$', s):
        return s  # نتركه كما هو — justify=right تكفي
    return s

def clean(val):
    """تنظيف القيمة عند الحفظ"""
    return str(val).replace("\u200f","").replace("\u200e","").strip()

def fmt_raqm(raqm):
    """تنسيق رقم الدورة: 13/2026 → 2026/13، 1/2026 → 2026/01"""
    import re as _re
    s = str(raqm or "").strip()
    if not s: return s
    m = _re.match(r'^(\d+)/(\d{4})$', s)
    if m: return f"{m.group(2)}/{m.group(1).zfill(2)}"
    parts = s.split("/")
    fmt_parts = []
    for p in parts:
        p = p.strip()
        if p.isdigit() and len(p) == 1:
            fmt_parts.append(f"0{p}")
        else:
            fmt_parts.append(p)
    return "/".join(fmt_parts)

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title("نظام إدارة شهادات الكفاءة المهنية — وزارة النقل")
        self.geometry("1280x780"); self.configure(bg=BG); self.resizable(True,True)
        init_excel(); self._build()
    def _build(self):
        hdr=tk.Frame(self,bg=BG3,height=60); hdr.pack(fill="x")
        tk.Label(hdr,text="🚌  نظام إدارة شهادات الكفاءة المهنية",font=_efont(15,bold=True),bg=BG3,fg=ACCENT2).pack(side="right",padx=20,pady=12)
        tk.Label(hdr,text="وزارة النقل — الجمهورية الجزائرية الديمقراطية الشعبية",font=("Arial",9),bg=BG3,fg=TEXT2).pack(side="left",padx=20)
        nav=tk.Frame(self,bg=BG2); nav.pack(fill="x")
        self.pages={}; self.btns={}
        for label,key in [
            ("🏢 المركز",       "markaz"),
            ("👤 المترشحون",    "mtrsh"),
            ("📅 الدورات",      "dawra"),
            ("📄 الشهادات",     "shahada"),
            ("📊 الإحصائيات",  "stats"),
            ("🗂 تسيير الدورة",  "gestion"),
        ]:
            b=tk.Button(nav,text=label,font=_efont(10,bold=True),bg=BG2,fg=TEXT2,relief="flat",padx=16,pady=11,cursor="hand2",command=lambda k=key:self._show(k))
            b.pack(side="right"); self.btns[key]=b
        self.content=tk.Frame(self,bg=BG); self.content.pack(fill="both",expand=True,padx=8,pady=8)
        self.pages["markaz"]=MarkazPage(self.content,self)
        self.pages["dawra"]=DawraPage(self.content,self)
        self.pages["mtrsh"]=MtrshinPage(self.content,self)
        self.pages["shahada"]=ShahadaPage(self.content,self)
        self.pages["stats"]=StatsPage(self.content,self)
        self.pages["gestion"]=DawraGestionPage(self.content,self)
        self._show("markaz")
    def _show(self, key):
        for k,p in self.pages.items(): p.pack_forget(); self.btns[k].configure(bg=BG2,fg=TEXT2)
        self.pages[key].pack(fill="both",expand=True); self.btns[key].configure(bg=ACCENT,fg=WHITE)
        # عند الانتقال لقسم المترشحين — نحمل رقم الدورة المحددة تلقائياً
        if key == "gestion":
            try:
                dawra_page = self.pages.get("dawra")
                if dawra_page:
                    sel = dawra_page.tree.selection()
                    if sel:
                        raqm = str(dawra_page.tree.item(sel[0])["values"][0])
                        self.pages["gestion"].raqm_var.set(raqm)
                        self.pages["gestion"]._load_dawra()
            except Exception: pass
        if key == "mtrsh":
            try:
                dawra_page = self.pages.get("dawra")
                if dawra_page:
                    sel = dawra_page.tree.selection()
                    if sel:
                        raqm = str(dawra_page.tree.item(sel[0])["values"][0])
                        self.pages["mtrsh"].mv["رقم الدورة"].set(raqm)
            except Exception: pass
        self.pages[key].refresh()

class MarkazPage(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=BG); self.app=app; self._build()
    def _build(self):
        tk.Label(self,text="🏢  معلومات مركز التكوين",font=_efont(14,bold=True),bg=BG,fg=ACCENT2).pack(pady=(8,2))
        tk.Label(self,text="تُدخل مرة واحدة وتُستخدم في جميع الوثائق",font=("Arial",9),bg=BG,fg=TEXT2).pack()
        frame=tk.Frame(self,bg=BG2,padx=30,pady=20); frame.pack(fill="both",expand=True,padx=20,pady=8)
        self.vars={}
        for i,col in enumerate(COLS_MARKAZ):
            r,c=divmod(i,2)
            tk.Label(frame,text=col+" :",font=("Arial",9,"bold"),bg=BG2,fg=TEXT2,anchor="e").grid(row=r,column=c*2,sticky="e",padx=8,pady=5)
            var=tk.StringVar()
            if col == "مديرية النقل لولاية":
                w=ttk.Combobox(frame,textvariable=var,values=WILAYAS,
                               width=32,font=("Arial",10),justify="right")
                w.grid(row=r,column=c*2+1,sticky="ew",padx=8,pady=5)
            else:
                tk.Entry(frame,textvariable=var,width=34,font=("Arial",10),bg=BG3,fg=TEXT,insertbackground=TEXT,relief="flat",bd=4,justify="right").grid(row=r,column=c*2+1,sticky="ew",padx=8,pady=5)
            self.vars[col]=var
        tk.Label(frame,text="آخر رقم ملف (للبدء منه):",bg=BG2,fg="#f5c842",
                 font=("Arial",9,"bold")).grid(row=7,column=0,sticky="e",pady=4,padx=8)
        var_milaf=tk.StringVar()
        tk.Entry(frame,textvariable=var_milaf,width=12,font=("Arial",10),
                 bg="#2e3450",fg="#f5c842",insertbackground=TEXT,relief="flat",
                 bd=4,justify="center").grid(row=7,column=1,sticky="w",padx=8,pady=4)
        self.vars["آخر رقم ملف"]=var_milaf
        tk.Label(frame,text="مثال: 2847 — سيبدأ البرنامج من 2848",
                 bg=BG2,fg=TEXT2,font=("Arial",7)).grid(row=7,column=2,sticky="w",padx=4)
        bf=tk.Frame(frame,bg=BG2); bf.grid(row=8,column=0,columnspan=4,pady=12)
        sbtn(bf,"‏💾 حفظ",self._save,ACCENT).pack(side="right",padx=5)
        sbtn(bf,"‏🔄 تحديث",self.refresh,BG3).pack(side="right",padx=5)
        sbtn(bf,"‏📦 نسخ احتياطي",self._backup,"#8e44ad").pack(side="right",padx=5)
        sbtn(bf,"‏📤 تصدير Excel",self._export_xl,"#16a085").pack(side="right",padx=5)
        sbtn(bf,"‏📥 استيراد Excel",self._import_xl,"#d35400").pack(side="right",padx=5)

    def _backup(self):
        if not _USE_DB: messagebox.showwarning("⚠️","SQLite غير متاح"); return
        try:
            ts = db.backup_db()
            backups = db.get_backup_list()
            messagebox.showinfo("✅ نسخ احتياطي",
                f"تم حفظ نسخة احتياطية ✅\n"
                f"التاريخ: {ts}\n"
                f"إجمالي النسخ: {len(backups)}")
        except Exception as e: messagebox.showerror("❌", f"خطأ: {e}")

    def _export_xl(self):
        path = filedialog.asksaveasfilename(
            title="تصدير Excel", defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile="بيانات_مركز_التكوين.xlsx")
        if not path: return
        try:
            if _USE_DB: db.export_to_excel(path)
            else:
                import shutil; shutil.copy2(EXCEL_FILE, path)
            messagebox.showinfo("✅", f"تم التصدير:\n{path}")
        except Exception as e: messagebox.showerror("❌", f"خطأ: {e}")

    def _import_xl(self):
        path = filedialog.askopenfilename(
            title="استيراد من Excel",
            filetypes=[("Excel","*.xlsx *.xls")])
        if not path: return
        if not messagebox.askyesno("استيراد",
            "سيتم استيراد البيانات من Excel لـ SQLite.\nهل تريد المتابعة؟"): return
        try:
            if _USE_DB:
                result = db.import_from_excel(path)
                messagebox.showinfo("✅",
                    f"تم الاستيراد:\n"
                    f"دورات: {result['دورات']}\n"
                    f"مترشحون: {result['مترشحون']}")
            else:
                import shutil; shutil.copy2(path, EXCEL_FILE)
                messagebox.showinfo("✅","تم نسخ الملف")
        except Exception as e: messagebox.showerror("❌", f"خطأ: {e}")

    def refresh(self):
        data = get_markaz()
        import re
        for i, col in enumerate(COLS_MARKAZ):
            val = str(data[i] if i < len(data) and data[i] else "")
            # تحويل DD/MM/YYYY → YYYY/MM/DD
            m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', val.strip())
            if m: val = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
            # تحويل رقم/سنة → سنة/رقم (مثل 442/2024 → 2024/442)
            m2 = re.match(r'^(\d+)/(\d{4})$', val.strip())
            if m2: val = f"{m2.group(2)}/{m2.group(1)}"
            self.vars[col].set(val)
    def _save(self):
        import re
        raw = []
        for c in COLS_MARKAZ:
            val = self.vars[c].get().replace("\u200f","").strip()
            # إعادة الصيغة الأصلية عند الحفظ: YYYY/XX → XX/YYYY
            m = re.match(r'^(\d{4})/(\d+)$', val)
            if m and not re.match(r'^\d{4}/\d{2}/\d{2}$', val):
                val = f"{m.group(2)}/{m.group(1)}"
            raw.append(val)
        data = raw
        wb=openpyxl.load_workbook(EXCEL_FILE); ws=wb[SHEETS["مركز"]]
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row: cell.value = None
        for i,v in enumerate(data,1):
            ws.cell(row=2,column=i,value=v)
        wb.save(EXCEL_FILE)
        # حفظ في SQLite
        if _USE_DB:
            try:
                d = dict(zip(COLS_MARKAZ, data))
                db.save_markaz(d)
            except Exception as e: print(f"SQLite markaz: {e}")
        # نسخ احتياطي تلقائي
        if _USE_DB:
            try: db.backup_db()
            except Exception: pass
        messagebox.showinfo("✅","تم حفظ معلومات المركز والنسخ الاحتياطي")

class DawraPage(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=BG); self.app=app; self._build()
    def _build(self):
        tk.Label(self,text="📅  إدارة الدورات التكوينية",font=_efont(14,bold=True),bg=BG,fg=ACCENT2).pack(pady=(8,2))
        top=tk.Frame(self,bg=BG); top.pack(fill="both",expand=True,padx=20,pady=5)
        form=tk.LabelFrame(top,text="  إضافة دورة  ",font=("Arial",10,"bold"),bg=BG2,fg=ACCENT,padx=15,pady=12)
        form.pack(side="right",fill="y"); self.dvars={}
        for i,col in enumerate(["رقم الدورة","تاريخ افتتاح الدورة","تاريخ اختتام الدورة"]):
            tk.Label(form,text=col+":",bg=BG2,fg=TEXT2,font=("Arial",9)).grid(row=i,column=0,sticky="e",pady=4,padx=5)
            var=tk.StringVar()
            tk.Entry(form,textvariable=var,width=20,bg=BG3,fg=TEXT,relief="flat",bd=3,justify="right",insertbackground=TEXT).grid(row=i,column=1,pady=4,padx=5)
            self.dvars[col]=var

        # رقم المحضر — مستقل عن رقم الدورة (row=4 بعد الخانات الثلاث)
        tk.Label(form,text="رقم المحضر:",bg=BG2,fg="#f5c842",font=("Arial",9,"bold")).grid(row=4,column=0,sticky="e",pady=4,padx=5)
        var_mahdar=tk.StringVar()
        tk.Entry(form,textvariable=var_mahdar,width=20,bg=BG3,fg="#f5c842",
                 relief="flat",bd=3,justify="right",insertbackground=TEXT,
                 font=("Arial",9,"bold")).grid(row=4,column=1,pady=4,padx=5)
        self.dvars["رقم المحضر"]=var_mahdar
        # تحويل تلقائي لصيغة YYYY/MM/DD عند فقدان التركيز
        def _auto_fmt_date(col_name, *a):
            val = self.dvars[col_name].get().strip()
            if not val: return
            import re
            m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', val)
            if m:
                self.dvars[col_name].set(f"{m.group(3)}/{m.group(2)}/{m.group(1)}")
            m2 = re.match(r'^(\d{4})-(\d{2})-(\d{2})$', val)
            if m2:
                self.dvars[col_name].set(f"{m2.group(1)}/{m2.group(2)}/{m2.group(3)}")
        for col in ["تاريخ افتتاح الدورة","تاريخ اختتام الدورة"]:
            self.dvars[col].trace_add("write", lambda *a, c=col: _auto_fmt_date(c))
        # تحويل رقم الدورة XX/YYYY → YYYY/XX
        def _auto_fmt_raqm(*a):
            val = self.dvars["رقم الدورة"].get().strip()
            if not val: return
            import re
            m = re.match(r'^(\d+)/(\d{4})$', val)
            if m:
                self.dvars["رقم الدورة"].set(f"{m.group(2)}/{m.group(1).zfill(2)}")
        self.dvars["رقم الدورة"].trace_add("write", _auto_fmt_raqm)
        tk.Label(form,text="التاريخ: YYYY/MM/DD",bg=BG2,fg=TEXT2,font=("Arial",7)).grid(row=5,column=0,columnspan=2)
        # ── أزرار الإضافة والتعديل والحذف ────────────────────
        bf_top=tk.Frame(form,bg=BG2); bf_top.grid(row=6,column=0,columnspan=2,pady=6,sticky="ew")
        sbtn(bf_top,"‏➕ إضافة",self._add,SUCCESS,pady=5).pack(side="right",padx=2,fill="x",expand=True)
        sbtn(bf_top,"‏✏️ تعديل",self._edit_dawra,WARNING,pady=5).pack(side="right",padx=2,fill="x",expand=True)
        sbtn(bf_top,"‏🗑️ حذف",self._del_dawra,DANGER,pady=5).pack(side="right",padx=2,fill="x",expand=True)
        # ── أزرار المحاضر ──────────────────────────────────────
        tk.Label(form,text="── محاضر الدورة ──",bg=BG2,fg=TEXT2,font=("Arial",7)).grid(row=5,column=0,columnspan=2,pady=(6,2))

        # ── إطار قابل للتمرير للأزرار ──────────────────────────
        scroll_frame = tk.Frame(form, bg=BG2)
        scroll_frame.grid(row=7, column=0, columnspan=2, pady=2, sticky="ew")
        _canvas = tk.Canvas(scroll_frame, bg=BG2, highlightthickness=0, width=220, height=280)
        _vsb = tk.Scrollbar(scroll_frame, orient="vertical", command=_canvas.yview)
        _canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side="left", fill="y")
        _canvas.pack(side="right", fill="both", expand=True)
        bf_m = tk.Frame(_canvas, bg=BG2)
        _win_id = _canvas.create_window((0,0), window=bf_m, anchor="nw")
        def _on_fc(e): _canvas.configure(scrollregion=_canvas.bbox("all"))
        def _on_cc(e): _canvas.itemconfig(_win_id, width=e.width)
        bf_m.bind("<Configure>", _on_fc)
        _canvas.bind("<Configure>", _on_cc)
        def _on_mw(e): _canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        _canvas.bind("<MouseWheel>", _on_mw)

        # ── المرحلة 1: تسجيل المترشحين ──────────────────────
        tk.Label(bf_m,text="① تسجيل المترشحين",bg=BG2,fg="#f5c842",
                 font=("Arial",8,"bold")).pack(fill="x",padx=4,pady=(6,1))
        sbtn(bf_m,"‏👤 إضافة مترشح",self._go_to_mtrsh,"#1a6b3c",pady=4).pack(fill="x",padx=4,pady=2)

        # ── المرحلة 2: فتح الدورة ────────────────────────────
        tk.Label(bf_m,text="② فتح الدورة",bg=BG2,fg="#f5c842",
                 font=("Arial",8,"bold")).pack(fill="x",padx=4,pady=(6,1))
        sbtn(bf_m,"‏📄 محضر الفتح",self._mahdar_fath,"#2471a3",pady=4).pack(fill="x",padx=4,pady=2)
        sbtn(bf_m,"‏📨 جدول إرسال الفتح",self._jadwal_irsal_fath,"#1a5276",pady=4).pack(fill="x",padx=4,pady=2)
        sbtn(bf_m,"‏📊 قائمة المترشحين",self._liste_mtrsheen,"#6c3483",pady=4).pack(fill="x",padx=4,pady=2)

        # ── المرحلة 3: النقاط ────────────────────────────────
        tk.Label(bf_m,text="③ الامتحان والنقاط",bg=BG2,fg="#f5c842",
                 font=("Arial",8,"bold")).pack(fill="x",padx=4,pady=(6,1))
        sbtn(bf_m,"‏🎲 نقاط تلقائية",self._gen_notes,"#784212",pady=4).pack(fill="x",padx=4,pady=2)
        sbtn(bf_m,"‏✏️ نقاط يدوية",self._manual_notes,"#4a235a",pady=4).pack(fill="x",padx=4,pady=2)

        # ── المرحلة 4: مداولة الدورة ─────────────────────────
        tk.Label(bf_m,text="④ مداولة الدورة",bg=BG2,fg="#f5c842",
                 font=("Arial",8,"bold")).pack(fill="x",padx=4,pady=(6,1))
        sbtn(bf_m,"‏📋 محضر المداولة",self._mahdar_mudalawa,"#117a65",pady=4).pack(fill="x",padx=4,pady=2)
        sbtn(bf_m,"‏📨 جدول إرسال المداولة",self._jadwal_irsal_mudalawa,"#0e6655",pady=4).pack(fill="x",padx=4,pady=2)
        sbtn(bf_m,"‏🏆 قائمة الناجحين + النقاط",self._liste_najihin,"#1a5276",pady=4).pack(fill="x",padx=4,pady=2)
        tbl=tk.Frame(top,bg=BG); tbl.pack(side="right",fill="both",expand=True,padx=(0,10))
        cols=["رقم الدورة","تاريخ الافتتاح","تاريخ الاختتام","عدد المترشحين","الحالة"]
        self.tree,sb=mk_tree(tbl,cols,14)
        for col,w in zip(cols,[120,130,130,120,100]):
            self.tree.heading(col,text=col,anchor="center")
            self.tree.column(col,width=w,minwidth=60,anchor="center",stretch=True)
        self.tree.pack(side="right",fill="both",expand=True); sb.pack(side="left",fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_dawra_select)

    def _on_dawra_select(self, event=None):
        """عند تحديد دورة — تحميل البيانات تلقائياً"""
        sel = self.tree.selection()
        if not sel: return
        raqm = str(self.tree.item(sel[0])["values"][0])
        for r in load_sheet("دورات"):
            if not any(r): continue
            if fmt_raqm(str(r[2] or "")) == raqm:
                mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                self.dvars["رقم المحضر"].set(mahdar)
                self.dvars["رقم الدورة"].set(raqm)
                self.dvars["تاريخ افتتاح الدورة"].set(fmt_date(str(r[0] or "")))
                self.dvars["تاريخ اختتام الدورة"].set(fmt_date(str(r[1] or "")))
                break

    def refresh(self):
        for r in self.tree.get_children(): self.tree.delete(r)
        # تلوين حسب الحالة
        self.tree.tag_configure("مفتوحة",  background="#1a3a1a", foreground="#2ecc71")
        self.tree.tag_configure("جارية",   background="#1a2a3a", foreground="#3498db")
        self.tree.tag_configure("منتهية",  background="#2a2a2a", foreground="#7f8c8d")
        for r in load_sheet("دورات"):
            if not any(r): continue
            n = len(r)
            d1    = fmt_date(str(r[0] or "")) if n > 0 else ""
            d2    = fmt_date(str(r[1] or "")) if n > 1 else ""
            raqm  = fmt_raqm(str(r[2] or "")) if n > 2 else ""
            cnt   = str(r[3] or "—")          if n > 3 else "—"
            hala  = str(r[10] or "مفتوحة")    if n > 10 else "مفتوحة"
            self.tree.insert("","end",values=(raqm, d1, d2, cnt, hala), tags=(hala,))

    def _edit_dawra(self):
        """تعديل بيانات دورة محددة"""
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر دورة أولاً"); return
        vals = self.tree.item(sel[0])["values"]
        raqm_old = str(vals[0])
        # ملء الفورم ببيانات الدورة المحددة
        self.dvars["رقم الدورة"].set(raqm_old)
        self.dvars["تاريخ افتتاح الدورة"].set(fmt_date(str(vals[1])))
        self.dvars["تاريخ اختتام الدورة"].set(fmt_date(str(vals[2])))
        # تحميل رقم المحضر من Excel
        for r in load_sheet("دورات"):
            if not any(r): continue
            if fmt_raqm(str(r[2] or "")) == raqm_old:
                mahdar = str(r[9] or "") if len(r) > 9 else ""
                if "رقم المحضر" in self.dvars:
                    self.dvars["رقم المحضر"].set(mahdar)
                break
        self._edit_mode = raqm_old
        messagebox.showinfo("✏️ تعديل",
            f"تم تحميل بيانات الدورة {raqm_old}\n"
            "عدّل البيانات ثم اضغط ✅ حفظ التعديل")
        # تغيير زر الإضافة مؤقتاً
        self._confirm_edit_btn = tk.Button(
            self.tree.master.master,
            text="✅ حفظ التعديل",
            command=self._confirm_edit_dawra,
            font=("Arial",10,"bold"), bg=SUCCESS, fg=WHITE,
            relief="flat", padx=10, pady=5, cursor="hand2"
        )
        self._confirm_edit_btn.pack(pady=4)

    def _confirm_edit_dawra(self):
        """تأكيد حفظ تعديل الدورة"""
        if not hasattr(self,'_edit_mode'): return
        raqm_old = self._edit_mode
        raqm_new = self.dvars["رقم الدورة"].get().strip()
        d1 = self.dvars["تاريخ افتتاح الدورة"].get().strip()
        d2 = self.dvars["تاريخ اختتام الدورة"].get().strip()
        if not raqm_new: messagebox.showwarning("⚠️","أدخل رقم الدورة"); return
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb[SHEETS["دورات"]]
            for row in ws.iter_rows(min_row=2):
                if str(row[2].value or "").strip() == str(raqm_old).strip():
                    row[0].value = d1
                    row[1].value = d2
                    row[2].value = raqm_new
                    break
            wb.save(EXCEL_FILE)
            for v in self.dvars.values(): v.set("")
            self._edit_mode = None
            if hasattr(self,'_confirm_edit_btn'):
                self._confirm_edit_btn.destroy()
            self.refresh()
            messagebox.showinfo("✅",f"تم تحديث الدورة {raqm_new}")
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح","أغلق Excel أولاً ثم حاول مرة أخرى")
        except Exception as e:
            messagebox.showerror("❌", f"خطأ:\n{e}")

    def _del_dawra(self):
        """حذف دورة محددة"""
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر دورة أولاً"); return
        raqm = str(self.tree.item(sel[0])["values"][0])
        if not messagebox.askyesno("🗑️ حذف",
            f"هل تريد حذف الدورة {raqm}؟\nتحذير: لن تُحذف بيانات المترشحين."): return
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb[SHEETS["دورات"]]
            for i, row in enumerate(ws.iter_rows(min_row=2), 2):
                if str(row[2].value or "").strip() == str(raqm).strip():
                    ws.delete_rows(i); break
            wb.save(EXCEL_FILE)
            self.refresh()
            messagebox.showinfo("✅",f"تم حذف الدورة {raqm}")
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح","أغلق Excel أولاً ثم حاول مرة أخرى")
        except Exception as e:
            messagebox.showerror("❌", f"خطأ:\n{e}")

    def _go_to_mtrsh(self):
        """الانتقال لقسم المترشحين مع تثبيت رقم الدورة"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("⚠️","اختر دورة أولاً"); return
        vals = self.tree.item(sel[0])["values"]
        raqm = str(vals[0])
        hala = str(vals[4]) if len(vals) > 4 else "مفتوحة"
        if hala != "مفتوحة":
            messagebox.showwarning("⚠️",
                f"الدورة {raqm} حالتها '{hala}'\n"
                "التسجيل مغلق — يُسمح فقط للدورات المفتوحة.")
            return
        mtrsh_page = self.app.pages.get("mtrsh")
        if mtrsh_page:
            mtrsh_page.mv["رقم الدورة"].set(raqm)
        self.app._show("mtrsh")

    def _get_raqm_dawra(self):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر دورة أولاً"); return None
        return str(self.tree.item(sel[0])["values"][0])

    def _mahdar_fath(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        mahdar = self.dvars["رقم المحضر"].get().strip() if "رقم المحضر" in self.dvars else ""
        if not mahdar:
            for r in load_sheet("دورات"):
                if not any(r): continue
                if fmt_raqm(str(r[2] or "")) == raqm:
                    mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                    break
        try:
            import محاضر_الدورة as md
            path = md.fill_mahdar_fath(raqm, raqm_mahdar=mahdar if mahdar else None)
            if path:
                # تغيير حالة الدورة → جارية
                _set_dawra_hala(raqm, "جارية")
                self.refresh()
                md.open_file(path)
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _jadwal_irsal_fath(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        mahdar = self.dvars["رقم المحضر"].get().strip() if "رقم المحضر" in self.dvars else ""
        if not mahdar:
            for r in load_sheet("دورات"):
                if not any(r): continue
                if fmt_raqm(str(r[2] or "")) == raqm:
                    mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                    break
        try:
            import محاضر_الدورة as md
            path = md.fill_jadwal_irsal_fath(raqm, raqm_mahdar=mahdar if mahdar else None)
            if path: md.open_file(path)
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _mahdar_mudalawa(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        mahdar = self.dvars["رقم المحضر"].get().strip() if "رقم المحضر" in self.dvars else ""
        if not mahdar:
            for r in load_sheet("دورات"):
                if not any(r): continue
                if fmt_raqm(str(r[2] or "")) == raqm:
                    mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                    break
        try:
            import محاضر_الدورة as md
            path = md.fill_mahdar_mudalawa(raqm, raqm_mahdar=mahdar if mahdar else None)
            if path:
                # تغيير حالة الدورة → منتهية
                _set_dawra_hala(raqm, "منتهية")
                self.refresh()
                md.open_file(path)
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _jadwal_irsal_mudalawa(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        mahdar = self.dvars["رقم المحضر"].get().strip() if "رقم المحضر" in self.dvars else ""
        if not mahdar:
            for r in load_sheet("دورات"):
                if not any(r): continue
                if fmt_raqm(str(r[2] or "")) == raqm:
                    mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                    break
        try:
            import محاضر_الدورة as md
            path = md.fill_jadwal_irsal_mudalawa(raqm, raqm_mahdar=mahdar if mahdar else None)
            if path: md.open_file(path)
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _liste_mtrsheen(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        try:
            import محاضر_الدورة as md
            out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "محاضر_PDF")
            os.makedirs(out_dir, exist_ok=True)
            path = md.export_liste_mtrsheen(raqm, out_dir)
            if path:
                if messagebox.askyesno("✅ تم", f"تم إنشاء قائمة المترشحين\nفتح الملف؟"):
                    md.open_file(path)
            else:
                messagebox.showwarning("⚠️", "لا يوجد مترشحون في هذه الدورة")
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _liste_najihin(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        try:
            import محاضر_الدورة as md
            out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "محاضر_PDF")
            os.makedirs(out_dir, exist_ok=True)
            path = md.export_liste_najihin(raqm, out_dir)
            if path:
                if messagebox.askyesno("✅ تم", f"تم إنشاء قائمة الناجحين\nفتح الملف؟"):
                    md.open_file(path)
            else:
                messagebox.showwarning("⚠️", "لا يوجد ناجحون في هذه الدورة")
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _gen_notes(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        if not messagebox.askyesno("تأكيد", f"توليد نقاط تلقائية لدورة {raqm}؟"): return
        try:
            import محاضر_الدورة as md
            mtrsheen = md.load_mtrsheen(raqm)
            if not mtrsheen: messagebox.showwarning("⚠️","لا يوجد مترشحون"); return
            notes_data = {}
            for m in mtrsheen:
                nom  = str(m.get("الاسم واللقب","") or "")
                takh = str(m.get("تخصص التكوين","") or "")
                n = md.gen_notes_auto(takh); n["التخصص"] = takh
                notes_data[nom] = n
            md.save_notes_to_excel(raqm, notes_data)

            # ── تحديث حالة النجاح في ورقة المترشحين ─────────
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb[SHEETS["مترشحون"]]
            rows = load_sheet("مترشحون")
            today = datetime.now().strftime("%Y-%m-%d")
            for i, r in enumerate(rows):
                nom = str(r[1] or "").strip()
                if nom in notes_data:
                    qar = notes_data[nom].get("القرار","")
                    ws.cell(row=i+2, column=20, value=qar)  # حالة النجاح
                    if qar == "ناجح":
                        if not (r[20] if len(r)>20 else ""):
                            ws.cell(row=i+2, column=21, value=next_file_num())  # رقم الشهادة
                            ws.cell(row=i+2, column=22, value=today)             # تاريخ الشهادة
                        takh_r = str(r[16] or "") if len(r)>16 else ""
                        if "البضائع" in takh_r and "الأشخاص" in takh_r:
                            if not (r[22] if len(r)>22 else ""):
                                ws.cell(row=i+2, column=23, value=next_file_num())  # رقم الشهادة 2
                                ws.cell(row=i+2, column=24, value=today)             # تاريخ الشهادة 2
            wb.save(EXCEL_FILE)
            messagebox.showinfo("✅", f"تم توليد نقاط {len(notes_data)} مترشح وتحديث الحالة")
            if messagebox.askyesno("📊 عرض النقاط", "هل تريد فتح ملف Excel لعرض النقاط؟"):
                md.open_file(md.EXCEL_FILE)
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح",
                "ملف Excel مفتوح حالياً!\n\n"
                "أغلق ملف 'بيانات_مركز_التكوين.xlsx' من Excel\n"
                "ثم حاول مرة أخرى.")
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

    def _manual_notes(self):
        raqm = self._get_raqm_dawra()
        if not raqm: return
        try:
            import محاضر_الدورة as md
            mtrsheen = md.load_mtrsheen(raqm)
            if not mtrsheen: messagebox.showwarning("⚠️","لا يوجد مترشحون"); return
        except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}"); return

        win = tk.Toplevel(); win.title("إدخال النقاط يدوياً"); win.configure(bg=BG)
        win.geometry("480x520"); win.grab_set()
        tk.Label(win, text=f"📝  نقاط الدورة : {raqm}",
                 font=("Arial",12,"bold"), bg=BG, fg=ACCENT2).pack(pady=8)
        f1 = tk.Frame(win, bg=BG); f1.pack(fill="x", padx=15, pady=4)
        tk.Label(f1, text="المترشح :", bg=BG, fg=TEXT2, font=("Arial",10)).pack(side="right", padx=6)
        names = [str(m.get("الاسم واللقب","")) for m in mtrsheen]
        v_nom = tk.StringVar(value=names[0] if names else "")
        cb = ttk.Combobox(f1, textvariable=v_nom, values=names, width=30, state="readonly", justify="right")
        cb.pack(side="right")
        notes_frame = tk.Frame(win, bg=BG2); notes_frame.pack(fill="both", expand=True, padx=15, pady=4)
        note_vars = {}
        entry_widgets = []

        def load_mtrsh_notes(*args):
            for w in notes_frame.winfo_children(): w.destroy()
            note_vars.clear()
            entry_widgets.clear()
            nom = v_nom.get()
            m = next((x for x in mtrsheen if str(x.get("الاسم واللقب","")) == nom), None)
            if not m: return
            takh = str(m.get("تخصص التكوين","") or "")
            cols, _ = md.get_mawad_list(takh)
            tk.Label(notes_frame, text=f"التخصص: {takh}", bg=BG2, fg=ACCENT2,
                     font=("Arial",9,"bold")).grid(row=0, column=0, columnspan=2, pady=4)
            for i, col in enumerate(cols, 1):
                tk.Label(notes_frame, text=f"{col} :", bg=BG2, fg=TEXT2,
                         font=("Arial",9)).grid(row=i, column=0, sticky="e", padx=8, pady=3)
                var = tk.StringVar(value="15")
                e = tk.Entry(notes_frame, textvariable=var, width=8, bg=BG3, fg=TEXT,
                             relief="flat", bd=3, justify="center",
                             font=("Arial",11))
                e.grid(row=i, column=1, pady=3)
                note_vars[col] = var
                entry_widgets.append(e)

            # ربط التنقل بالمفاتيح
            def make_nav(idx):
                def on_key(e):
                    if e.keysym in ("Return","Down"):
                        nxt = idx + 1
                        if nxt < len(entry_widgets):
                            entry_widgets[nxt].focus_set()
                            entry_widgets[nxt].select_range(0,"end")
                        else:
                            do_save()  # Enter في الأخير → حفظ
                        return "break"
                    elif e.keysym == "Up":
                        prv = idx - 1
                        if prv >= 0:
                            entry_widgets[prv].focus_set()
                            entry_widgets[prv].select_range(0,"end")
                        return "break"
                return on_key

            for idx, e in enumerate(entry_widgets):
                e.bind("<Return>", make_nav(idx))
                e.bind("<Down>",   make_nav(idx))
                e.bind("<Up>",     make_nav(idx))

            # التركيز على أول حقل
            if entry_widgets:
                entry_widgets[0].focus_set()
                entry_widgets[0].select_range(0,"end")

        v_nom.trace_add("write", load_mtrsh_notes); load_mtrsh_notes()

        def do_save():
            nom = v_nom.get()
            m = next((x for x in mtrsheen if str(x.get("الاسم واللقب","")) == nom), None)
            if not m: return
            takh = str(m.get("تخصص التكوين","") or "")
            try:
                notes = {col: float(var.get()) for col, var in note_vars.items()}
                moy = md.calc_moyenne(notes, takh)
                notes.update({"المعدل": moy, "القرار": "ناجح" if moy >= 10 else "راسب", "التخصص": takh})
                existing = {}
                try:
                    import openpyxl as ox
                    wb_n = ox.load_workbook(md.EXCEL_FILE)
                    if "النقاط" in wb_n.sheetnames:
                        ws_n = wb_n["النقاط"]
                        headers_n = [c.value for c in ws_n[1]]
                        for row in ws_n.iter_rows(min_row=2, values_only=True):
                            d = dict(zip(headers_n, row))
                            if str(d.get("رقم الدورة","")).strip() == str(raqm).strip():
                                existing[str(d.get("الاسم واللقب",""))] = d
                except Exception: pass
                existing[nom] = {"رقم الدورة": raqm, "الاسم واللقب": nom, **notes}
                md.save_notes_to_excel(raqm, existing)
                messagebox.showinfo("✅", f"تم حفظ نقاط {nom}\nالمعدل: {moy} — {notes['القرار']}")
            except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

        bf = tk.Frame(win, bg=BG); bf.pack(pady=8)
        tk.Button(bf, text="💾 حفظ", command=do_save, font=_efont(10,bold=True),
                  bg=SUCCESS, fg=WHITE, relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right", padx=6)
        tk.Button(bf, text="❌ إغلاق", command=win.destroy, font=("Arial",10),
                  bg=BG3, fg=TEXT, relief="flat", padx=10, pady=6, cursor="hand2").pack(side="right", padx=6)

    def _add(self):
        raqm   = self.dvars["رقم الدورة"].get().strip()
        d1     = self.dvars["تاريخ افتتاح الدورة"].get().strip()
        d2     = self.dvars["تاريخ اختتام الدورة"].get().strip()
        mahdar = self.dvars["رقم المحضر"].get().strip() if "رقم المحضر" in self.dvars else ""
        if not raqm or not d1: messagebox.showwarning("⚠️","أدخل رقم الدورة وتاريخ الافتتاح"); return
        save_row("دورات",[d1,d2,raqm,"","","","","","",mahdar,"مفتوحة"])  # حالة الدورة = مفتوحة
        for v in self.dvars.values(): v.set("")
        self.refresh(); messagebox.showinfo("✅","تمت إضافة الدورة رقم "+raqm)

class MtrshinPage(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=BG); self.app=app; self._build()
    def _build(self):
        tk.Label(self,text="👤  إدارة المترشحين",font=_efont(14,bold=True),bg=BG,fg=ACCENT2).pack(pady=(8,2))
        sb_f=tk.Frame(self,bg=BG2,pady=6); sb_f.pack(fill="x",padx=20)
        tk.Label(sb_f,text="🔍",bg=BG2,fg=TEXT).pack(side="right",padx=6)
        self.sv=tk.StringVar(); self.sv.trace_add("write",lambda *a:self.refresh())
        tk.Entry(sb_f,textvariable=self.sv,width=28,bg=BG3,fg=TEXT,relief="flat",bd=3,insertbackground=TEXT,justify="right").pack(side="right",padx=4)
        sbtn(sb_f,"‏📥 استيراد Excel",self._import,"#16a085").pack(side="left",padx=6)
        sbtn(sb_f,"‏📤 تصدير قالب",self._export_tpl,BG3).pack(side="left",padx=4)
        sbtn(sb_f,"‏📢 إرسال جماعي",self._send_bulk,"#8e44ad").pack(side="left",padx=4)
        sbtn(sb_f,"‏🔄 ترحيل البيانات",self._migrate_data,"#e67e22").pack(side="left",padx=4)
        sbtn(sb_f,"‏⚙️ الخانات",self._toggle_fields,"#5d6d7e").pack(side="left",padx=4)
        main=tk.Frame(self,bg=BG); main.pack(fill="both",expand=True,padx=20,pady=4)
        tbl=tk.Frame(main,bg=BG); tbl.pack(side="right",fill="both",expand=True)
        cols=["الاسم واللقب","تاريخ ومكان الازدياد","التخصص","رقم الدورة","الحالة","رقم الملف"]
        self.tree,scrollb=mk_tree(tbl,cols,14)
        for col,w in zip(cols,[200,180,120,90,80,120]):
            self.tree.heading(col,text=col,anchor="center")
            self.tree.column(col,width=w,minwidth=50,anchor="center",stretch=True)
        self.tree.tag_configure("najih",  background="#1a3a2a", foreground="#2ecc71")
        self.tree.tag_configure("rasib",  background="#3a1a1a", foreground="#e74c3c")
        self.tree.tag_configure("ghaib",  background="#2a2a2a", foreground="#95a5a6")
        self.tree.tag_configure("wajjal", background="#2a2000", foreground="#f39c12")
        self.tree.tag_configure("takhali",background="#3a1a00", foreground="#e67e22")
        self.tree.tag_configure("muid",   background="#001a3a", foreground="#3498db")
        self.tree.tag_configure("normal", background=BG2,       foreground=TEXT)
        self.tree.pack(side="right",fill="both",expand=True); scrollb.pack(side="left",fill="y")
        # عند اختيار مترشح من الجدول — تحميل بياناته في الفورم
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        # ── فورم قابل للتمرير ─────────────────────────────────────
        form_outer=tk.LabelFrame(main,text="  مترشح جديد  ",font=("Arial",9,"bold"),bg=BG2,fg=ACCENT,width=300)
        form_outer.pack(side="left",fill="y",padx=(8,0)); form_outer.pack_propagate(False)
        canvas_f=tk.Canvas(form_outer,bg=BG2,highlightthickness=0,width=280)
        scr_f=tk.Scrollbar(form_outer,orient="vertical",command=canvas_f.yview)
        canvas_f.configure(yscrollcommand=scr_f.set)
        scr_f.pack(side="right",fill="y"); canvas_f.pack(side="left",fill="both",expand=True)
        form=tk.Frame(canvas_f,bg=BG2,padx=6,pady=6)
        form_win=canvas_f.create_window((0,0),window=form,anchor="nw")
        def _on_frame_conf(e): canvas_f.configure(scrollregion=canvas_f.bbox("all"))
        def _on_canvas_conf(e): canvas_f.itemconfig(form_win,width=e.width)
        form.bind("<Configure>",_on_frame_conf)
        canvas_f.bind("<Configure>",_on_canvas_conf)
        def _on_mousewheel(e): canvas_f.yview_scroll(int(-1*(e.delta/120)),"units")
        canvas_f.bind_all("<MouseWheel>",_on_mousewheel)
        form.columnconfigure(0,weight=1)
        # ─────────────────────────────────────────────────────────
        self.mv={}
        # نوع الحقل: "ar"=عربي، "fr"=فرنسي كبير، "num"=أرقام، "nin"=18 رقم، "date"=تاريخ، "e"=عادي، "c"=قائمة
        ff=[
            ("الاسم واللقب",              "ar"),
            ("اللقب بالفرنسية",           "fr"),
            ("الاسم بالفرنسية",           "fr"),
            ("تاريخ ومكان الازدياد",      "ar"),
            ("العنوان",                   "ar"),
            ("رقم التعريف الوطني",        "e"),
            ("رقم رخصة السياقة",          "e"),
            ("تاريخ صدور رخصة السياقة",   "date"),
            ("تاريخ انتهاء رخصة السياقة", "date"),
            ("مكان إصدار الرخصة",         "ar"),
            ("فئات الرخصة",               "e"),
            ("زمرة الدم",                 "e"),
            ("الجنس",                     "e"),
            ("رقم هاتف المتربص",          "num"),
            ("البريد الإلكتروني",         "e"),
            ("تخصص التكوين",              "c"),
            ("رقم الدورة",                "e"),
            ("نوع التكوين",               "t"),  # t = combobox ثنائي
        ]
        self.ff = ff

        # ── تحميل إعدادات الإخفاء ────────────────────────────
        import json as _json
        cfg_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fields_config.json")

        # الخانات المخفية افتراضياً
        DEFAULT_HIDDEN = {
            "اللقب بالفرنسية",
            "الاسم بالفرنسية",
            "تاريخ انتهاء رخصة السياقة",
            "مكان إصدار الرخصة",
            "زمرة الدم",
            "البريد الإلكتروني",
            "الجنس",
        }
        try:
            with open(cfg_file, 'r', encoding='utf-8') as f:
                saved = _json.load(f)
            # إذا كان الملف يحتوي مفتاح "hidden" نستخدمه
            # وإلا نستخدم الافتراضي
            if "hidden" in saved:
                hidden = set(saved["hidden"])
            else:
                hidden = DEFAULT_HIDDEN
        except Exception:
            # أول تشغيل — نحفظ الافتراضي
            hidden = DEFAULT_HIDDEN
            try:
                with open(cfg_file, 'w', encoding='utf-8') as f:
                    _json.dump({"hidden": list(DEFAULT_HIDDEN)}, f, ensure_ascii=False)
            except Exception: pass
        self._hidden_fields = hidden
        self._cfg_file = cfg_file

        # ── دوال التحقق ── تُعرَّف مرة واحدة خارج الحلقة
        def _validate_num_fn(P):
            return P == "" or P.isdigit()
        def _validate_date_fn(P):
            return all(c.isdigit() or c == "/" for c in P)
        vcmd_num  = form.register(_validate_num_fn)
        vcmd_date = form.register(_validate_date_fn)

        self.entry_widgets = {}
        self.field_rows = {}

        for i,(col,wt) in enumerate(ff):
            row_idx = i*2
            lbl = tk.Label(form, text=col+" :", bg=BG2, fg=TEXT2,
                          font=("Arial",8), anchor="e")
            lbl.grid(row=row_idx, column=0, sticky="ew", pady=(4,0), padx=4)
            var = tk.StringVar()

            if wt=="t":  # نوع التكوين
                w = ttk.Combobox(form, textvariable=var, values=["أولي","متواصل"],
                                state="readonly", justify="right", font=("Arial",9))
                w.set("أولي")
            elif wt=="c":
                w = ttk.Combobox(form, textvariable=var, values=TAKHASSAS,
                                state="readonly", justify="right", font=("Arial",9))
                w.set(TAKHASSAS[0])
            else:
                bd_color = {"ar":"#2471a3","fr":"#117a65","nin":"#c0392b",
                           "num":"#784212","date":"#8e44ad"}.get(wt, BG3)
                w = tk.Entry(form, textvariable=var, bg=BG3, fg=TEXT,
                            relief="flat", bd=3, insertbackground=TEXT,
                            justify="right", font=("Arial",9),
                            highlightthickness=1,
                            highlightbackground=bd_color,
                            highlightcolor=bd_color)

                if wt == "ar":
                    def _set_ar(e, _w=w):
                        try:
                            import subprocess
                            subprocess.Popen(['powershell','-Command',
                                'Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.InputLanguage]::CurrentInputLanguage = [System.Windows.Forms.InputLanguage]::FromCulture([System.Globalization.CultureInfo]::GetCultureInfo("ar-DZ"))'],
                                creationflags=0x08000000)
                        except Exception: pass
                    w.bind("<FocusIn>", _set_ar)

                elif wt == "fr":
                    def _to_upper(e, _var=var):
                        val = _var.get()
                        up = val.upper()
                        if val != up:
                            _var.set(up)
                            e.widget.icursor(tk.END)
                    w.bind("<KeyRelease>", _to_upper)

                elif wt in ("num",):
                    w.config(validate="focusout", validatecommand=(vcmd_num, "%P"))

                elif wt == "date":
                    w.config(validate="focusout", validatecommand=(vcmd_date, "%P"))

                if wt == "nin":
                    lbl_nin = tk.Label(form, text="", bg=BG2,
                                      font=("Arial",7), fg=TEXT2)
                    lbl_nin.grid(row=row_idx+1, column=0, sticky="ew", padx=4)
                    def _check_nin(e, _var=var, _lbl=lbl_nin):
                        n = len(_var.get().strip())
                        if n == 0: _lbl.config(text="", fg=TEXT2)
                        elif n == 18: _lbl.config(text="✅ 18 رقم", fg="#2ecc71")
                        elif n < 18: _lbl.config(text=f"⚠️ {n}/18 رقم", fg="#f39c12")
                        else: _lbl.config(text=f"❌ {n} رقم — الحد 18", fg="#e74c3c")
                    w.bind("<KeyRelease>", _check_nin)

            w.grid(row=row_idx+1, column=0, sticky="ew", pady=(0,2), padx=4)
            self.mv[col] = var
            self.entry_widgets[col] = w
            self.field_rows[col] = (lbl, w)

            # إخفاء إذا كان في القائمة
            if col in self._hidden_fields:
                lbl.grid_remove()
                w.grid_remove()
        n=len(ff)*2
        # ── زر مسح رخصة السياقة ──────────────────────────────────
        sbtn(form,"‏📷 مسح رخصة السياقة تلقائياً",self._scan_license,
             "#1a6b3c",pady=6).grid(row=n,column=0,sticky="ew",padx=4,pady=(8,4))
        # صورة الرخصة المصغرة
        self.lbl_img = tk.Label(form,bg=BG2,text="لا توجد صورة",
                                fg=TEXT2,font=("Arial",8))
        self.lbl_img.grid(row=n+1,column=0,sticky="ew",padx=4,pady=2)
        n += 2
        bf=tk.Frame(form,bg=BG2); bf.grid(row=n,column=0,sticky="ew",pady=6)
        sbtn(bf,"‏➕ إضافة",self._add,SUCCESS,pady=5).pack(side="right",padx=2,fill="x",expand=True)
        sbtn(bf,"‏✏️ تحميل",self._edit,WARNING,pady=5).pack(side="right",padx=2)
        sbtn(bf,"‏💾 حفظ",self._save_edit,"#8e44ad",pady=5).pack(side="right",padx=2)
        tk.Label(form,text="── رسائل ──",bg=BG2,fg=TEXT2,font=("Arial",7)).grid(row=n+1,column=0,sticky="ew")
        bf3=tk.Frame(form,bg=BG2); bf3.grid(row=n+2,column=0,sticky="ew",pady=2)
        sbtn(bf3,"‏💬 واتساب",self._whatsapp,"#25D366",pady=5).pack(side="right",padx=2,fill="x",expand=True)
        sbtn(bf3,"‏✉️ إيميل",self._email,"#c0392b",pady=5).pack(side="right",padx=2,fill="x",expand=True)
        sbtn(bf3,"‏📱 SMS",self._sms,"#e67e22",pady=5).pack(side="right",padx=2,fill="x",expand=True)
        # ── شريط طباعة أسفل الصفحة ───────────────────────────────
        bar=tk.Frame(self,bg=BG3,pady=4); bar.pack(fill="x",padx=20,pady=(0,6),side="bottom")
        tk.Label(bar,text="🖨️ طباعة الوثائق :",bg=BG3,fg=TEXT2,font=_efont(8,bold=True)).pack(side="right",padx=8)
        sbtn(bar,"‏📝 Word تسجيل+وصل",self._print_word,"#16a085",pady=3).pack(side="right",padx=3)
        sbtn(bar,"‏📋 عقد التكوين",self._print_aqd,"#8e44ad",pady=3).pack(side="right",padx=3)
    def refresh(self):
        # ملء رقم الدورة المفتوحة تلقائياً إذا كان الحقل فارغاً
        if "رقم الدورة" in self.mv and not self.mv["رقم الدورة"].get().strip():
            raqm_maftuha = _get_dawra_maftuha()
            if raqm_maftuha:
                self.mv["رقم الدورة"].set(raqm_maftuha)
        for r in self.tree.get_children(): self.tree.delete(r)
        q=self.sv.get().strip().lower()
        for r in load_sheet("مترشحون"):
            if q and q not in str(r[1] or "").lower(): continue
            n = len(r)
            if n <= 16:
                dob    = fmt_date(str(r[2] or "")) if n > 2  else ""
                takh   = str(r[7] or "")           if n > 7  else ""
                dawra  = fmt_raqm(str(r[10] or ""))if n > 10 else ""
                status = str(r[11] or "")           if n > 11 else ""
                milaf  = str(r[12] or "")           if n > 12 else ""
            else:
                dob    = fmt_date(str(r[4]  or "")) if n > 4  else ""
                takh   = str(r[16] or "")           if n > 16 else ""
                dawra  = fmt_raqm(str(r[17] or "")) if n > 17 else ""
                status = str(r[19] or "")           if n > 18 else ""
                milaf  = str(r[26] or "")           if n > 25 else ""
            _tag_map = {"ناجح":"najih","راسب":"rasib","غائب":"ghaib",
                        "مؤجل":"wajjal","متخلي":"takhali","معيد":"muid"}
            tag = _tag_map.get(status, "normal")
            self.tree.insert("","end",values=(r[1],dob,takh,dawra,status or "—",milaf or "—"),tags=(tag,))

    def _toggle_fields(self):
        """نافذة إظهار/إخفاء الخانات — تعمل فوراً"""
        import json as _json
        win = tk.Toplevel()
        win.title("⚙️ الخانات")
        win.configure(bg=BG)
        win.geometry("280x500")
        win.grab_set()
        win.resizable(False, False)

        tk.Label(win, text="✅ ظاهر  |  ☐ مخفي",
                 bg=BG, fg=ACCENT2, font=("Arial",10,"bold")).pack(pady=8)

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=12)

        checks = {}
        for col, _ in self.ff:
            var = tk.BooleanVar(value=(col not in self._hidden_fields))

            def _on_change(c=col, v=var):
                """تطبيق فوري عند تغيير الحالة"""
                # البحث عن اللقب والخانة
                if c not in self.field_rows: return
                lbl, w = self.field_rows[c]
                if v.get():
                    # إظهار فوري
                    if lbl:
                        lbl.grid()
                    w.grid()
                    self._hidden_fields.discard(c)
                else:
                    # إخفاء فوري
                    if lbl:
                        lbl.grid_remove()
                    w.grid_remove()
                    self._hidden_fields.add(c)
                # حفظ الإعداد فوراً
                try:
                    with open(self._cfg_file, 'w', encoding='utf-8') as f:
                        _json.dump({"hidden": list(self._hidden_fields)},
                                   f, ensure_ascii=False)
                except Exception: pass

            cb = tk.Checkbutton(frame, text=col, variable=var,
                               command=_on_change,
                               bg=BG, fg=TEXT, selectcolor=BG3,
                               font=("Arial",10), anchor="e",
                               activebackground=BG, cursor="hand2")
            cb.pack(fill="x", pady=2)
            checks[col] = var

        tk.Button(win, text="✅ إغلاق", command=win.destroy,
                  bg=ACCENT, fg=WHITE, font=("Arial",10,"bold"),
                  relief="flat", padx=20, pady=6).pack(pady=10)

    def _on_tree_select(self, event=None):
        """عند اختيار مترشح من الجدول — تحميل بياناته في الفورم"""
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])["values"]
        if not vals: return

        nom = str(vals[0])  # الاسم واللقب
        # البحث في Excel عن المترشح
        for r in load_sheet("مترشحون"):
            if not any(r): continue
            n = len(r)
            r_nom = str(r[1] or "") if n > 1 else ""
            if r_nom.strip() == nom.strip():
                # خريطة الأعمدة
                mp = [
                    ("الاسم واللقب",              1),
                    ("اللقب بالفرنسية",           2),
                    ("الاسم بالفرنسية",           3),
                    ("تاريخ ومكان الازدياد",      4),
                    ("العنوان",                   5),
                    ("رقم التعريف الوطني",        6),
                    ("رقم رخصة السياقة",          7),
                    ("تاريخ صدور رخصة السياقة",   8),
                    ("تاريخ انتهاء رخصة السياقة", 9),
                    ("مكان إصدار الرخصة",         10),
                    ("فئات الرخصة",               11),
                    ("زمرة الدم",                 12),
                    ("الجنس",                     13),
                    ("رقم هاتف المتربص",          14),
                    ("البريد الإلكتروني",         15),
                    ("تخصص التكوين",              16),
                    ("رقم الدورة",                17),
                    ("نوع التكوين",               18),
                ]
                for col, cidx in mp:
                    val = str(r[cidx] or "") if cidx < n else ""
                    if col in ["تاريخ صدور رخصة السياقة","تاريخ انتهاء رخصة السياقة"]:
                        val = fmt_date(val)
                    if col == "رقم الدورة":
                        val = fmt_raqm(val)
                    if col == "نوع التكوين" and not val:
                        val = "أولي"
                    if col in self.mv:
                        self.mv[col].set(val)
                break

    def _scan_license(self):
        """مسح رخصة السياقة باستخدام Gemini AI"""
        import importlib.util, sys
        reader_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "قارئ_الرخصة.py")
        try:
            spec = importlib.util.spec_from_file_location("قارئ_الرخصة", reader_path)
            reader = importlib.util.module_from_spec(spec)
            sys.modules["قارئ_الرخصة"] = reader
            spec.loader.exec_module(reader)
            # نمرر self كـ parent_app حتى تُملأ الخانات تلقائياً
            reader.open_reader(parent_app=self)
        except FileNotFoundError:
            messagebox.showerror("❌", "ملف قارئ_الرخصة.py غير موجود")
        except Exception as e:
            messagebox.showerror("❌", f"خطأ:\n{e}")
    def _load_license_image(self, nin):
        try:
            import قراءة_الرخصة as ql
            path = ql.get_document_image(nin,"رخصة")
            if path:
                from PIL import Image, ImageTk
                img=Image.open(path); img.thumbnail((260,80))
                ph=ImageTk.PhotoImage(img)
                self.lbl_img.configure(image=ph,text=""); self.lbl_img.image=ph
            else:
                self.lbl_img.configure(image="",text="لا توجد صورة رخصة")
        except Exception: self.lbl_img.configure(text="")

    def _migrate_data(self):
        """ترحيل البيانات القديمة للهيكل الجديد"""
        if not messagebox.askyesno("🔄 ترحيل البيانات",
            "هذا الأمر سيرحّل بيانات المترشحين من الهيكل القديم للهيكل الجديد.\n\n"
            "يُستخدم مرة واحدة فقط عند الترقية.\n\n"
            "هل تريد المتابعة؟"):
            return
        try:
            import محاضر_الدورة as md
            n = md.migrate_mtrsheen()
            self.refresh()
            if n > 0:
                messagebox.showinfo("✅", f"تم ترحيل {n} مترشح بنجاح!\nيمكنك الآن استخدام كل الميزات.")
            else:
                messagebox.showinfo("✅", "البيانات محدَّثة مسبقاً — لا حاجة للترحيل.")
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح","أغلق Excel أولاً ثم حاول مرة أخرى.")
        except Exception as e:
            messagebox.showerror("❌", f"خطأ:\n{e}")

    def _add(self):
        nom=self.mv["الاسم واللقب"].get().strip()
        if not nom: messagebox.showwarning("⚠️","أدخل الاسم واللقب"); return

        # التحقق من رقم التعريف الوطني
        nin = self.mv["رقم التعريف الوطني"].get().strip()
        if nin and len(nin) != 18:
            messagebox.showerror("❌ رقم التعريف الوطني",
                f"رقم التعريف الوطني يجب أن يكون 18 رقماً بالضبط.\n"
                f"أدخلت {len(nin)} رقم.")
            return

        # ── التحقق من التكرار ─────────────────────────────────────────
        if nin:
            takh    = self.mv["تخصص التكوين"].get().strip()
            naw     = self.mv["نوع التكوين"].get().strip() or "أولي"
            raqm_d  = fmt_raqm(self.mv["رقم الدورة"].get().strip())
            sujuls  = [r for r in load_sheet("مترشحون")
                       if any(r) and str(r[6] if len(r)>6 else "").strip() == nin]

            for r in sujuls:
                n = len(r)
                r_takh   = str(r[16] if n>16 else "").strip()
                r_naw    = str(r[18] if n>18 else "").strip()
                r_dawra  = fmt_raqm(str(r[17] if n>17 else "").strip())
                r_shahada= str(r[21] if n>21 else "").strip()  # تاريخ الشهادة
                r_status = str(r[19] if n>19 else "").strip()  # حالة النجاح

                # نفس الدورة → تكرار مباشر
                if r_dawra and r_dawra == raqm_d:
                    messagebox.showerror("❌ مترشح مكرر",
                        f"هذا المترشح مسجل مسبقاً في نفس الدورة {raqm_d}\n"
                        f"رقم الملف: {r[26] if n>26 else '؟'}")
                    return

                # نفس التخصص + شهادة سارية
                if r_takh == takh and r_shahada:
                    from datetime import datetime as _dt
                    try:
                        exp = calc_exp(r_shahada)
                        exp_date = _dt.strptime(exp, "%Y/%m/%d")
                        if exp_date > _dt.now():
                            messagebox.showerror("❌ شهادة سارية",
                                f"هذا المترشح لديه شهادة سارية في تخصص '{takh}'\n"
                                f"تنتهي بتاريخ: {exp}\n\n"
                                f"لا يمكن تسجيله في نفس التخصص قبل انتهاء الشهادة.")
                            return
                    except Exception:
                        pass

                # نفس التخصص + لا شهادة + لا يزال في دورة جارية (حالة فارغة)
                if r_takh == takh and not r_shahada and r_status not in ("متخلي","غائب","راسب","مؤجل","معيد"):
                    if r_status == "" or r_status == "ناجح":
                        pass  # ناجح بدون شهادة نادر — نسمح
                    else:
                        res = messagebox.askyesno("⚠️ تحذير",
                            f"هذا المترشح مسجل مسبقاً في تخصص '{takh}' "
                            f"في الدورة {r_dawra}\n"
                            f"الحالة: {r_status or 'غير محددة'}\n\n"
                            f"هل تريد التسجيل رغم ذلك؟")
                        if not res: return

        # التحقق من رقم الهاتف
        tel = self.mv["رقم هاتف المتربص"].get().strip()
        if tel and not tel.isdigit():
            messagebox.showerror("❌ رقم الهاتف","رقم الهاتف يجب أن يحتوي أرقاماً فقط"); return

        # توليد رقم الملف
        raqm_milaf = next_file_num()

        data=[
            datetime.now().strftime("%Y-%m-%d"),        # 0
            self.mv["الاسم واللقب"].get(),               # 1
            self.mv["اللقب بالفرنسية"].get(),            # 2
            self.mv["الاسم بالفرنسية"].get(),            # 3
            self.mv["تاريخ ومكان الازدياد"].get(),       # 4
            self.mv["العنوان"].get(),                    # 5
            self.mv["رقم التعريف الوطني"].get(),         # 6
            self.mv["رقم رخصة السياقة"].get(),           # 7
            self.mv["تاريخ صدور رخصة السياقة"].get(),    # 8
            self.mv["تاريخ انتهاء رخصة السياقة"].get(),  # 9
            self.mv["مكان إصدار الرخصة"].get(),          # 10
            self.mv["فئات الرخصة"].get(),                # 11
            self.mv["زمرة الدم"].get(),                  # 12
            self.mv["الجنس"].get(),                      # 13
            self.mv["رقم هاتف المتربص"].get(),           # 14
            self.mv["البريد الإلكتروني"].get(),           # 15
            self.mv["تخصص التكوين"].get(),               # 16
            fmt_raqm(self.mv["رقم الدورة"].get()),       # 17 — مُطبَّع دائماً
            self.mv["نوع التكوين"].get() or "أولي",      # 18 نوع التكوين
            "","","","","","","",                        # 19-25
            raqm_milaf,                                  # 26 رقم الملف
        ]
        save_row("مترشحون", data)
        # بناء قاموس المترشح مع رقم الملف
        mtrsh_data = {
            "الاسم واللقب":                self.mv["الاسم واللقب"].get(),
            "اللقب بالفرنسية":             self.mv["اللقب بالفرنسية"].get(),
            "الاسم بالفرنسية":             self.mv["الاسم بالفرنسية"].get(),
            "تاريخ ومكان الازدياد":        self.mv["تاريخ ومكان الازدياد"].get(),
            "العنوان":                     self.mv["العنوان"].get(),
            "رقم التعريف الوطني":          self.mv["رقم التعريف الوطني"].get(),
            "رقم رخصة السياقة":            self.mv["رقم رخصة السياقة"].get(),
            "تاريخ صدور رخصة السياقة":     self.mv["تاريخ صدور رخصة السياقة"].get(),
            "تاريخ انتهاء رخصة السياقة":   self.mv["تاريخ انتهاء رخصة السياقة"].get(),
            "مكان إصدار الرخصة":           self.mv["مكان إصدار الرخصة"].get(),
            "فئات الرخصة":                 self.mv["فئات الرخصة"].get(),
            "زمرة الدم":                   self.mv["زمرة الدم"].get(),
            "الجنس":                       self.mv["الجنس"].get(),
            "رقم هاتف المتربص":            self.mv["رقم هاتف المتربص"].get(),
            "البريد الإلكتروني":            self.mv["البريد الإلكتروني"].get(),
            "تخصص التكوين":               self.mv["تخصص التكوين"].get(),
            "رقم الدورة":                  fmt_raqm(self.mv["رقم الدورة"].get()),
            "نوع التكوين":                 self.mv["نوع التكوين"].get() or "أولي",
            "تاريخ التسجيل":               datetime.now().strftime("%Y-%m-%d"),
            "رقم الملف":                   raqm_milaf,
        }
        for v in self.mv.values(): v.set("")
        self.mv["تخصص التكوين"].set(TAKHASSAS[0])
        self.mv["نوع التكوين"].set("أولي")
        # ملء رقم الدورة المفتوحة تلقائياً للمترشح التالي
        raqm_maftuha = _get_dawra_maftuha()
        if raqm_maftuha:
            self.mv["رقم الدورة"].set(raqm_maftuha)
        self.refresh()
        # نافذة طباعة وثائق التسجيل
        self._show_print_dialog(mtrsh_data, nom)

    def _show_print_dialog(self, mtrsh_data, nom):
        """نافذة طباعة وثائق التسجيل بعد إضافة مترشح"""
        win = tk.Toplevel(); win.title("وثائق التسجيل"); win.configure(bg=BG)
        win.geometry("380x320"); win.grab_set()
        win.resizable(False, False)

        # رأس النافذة
        header = tk.Frame(win, bg="#1e2130", pady=12)
        header.pack(fill="x")
        tk.Label(header, text="✅  تمت إضافة المترشح بنجاح",
                 font=("Arial",12,"bold"), bg="#1e2130", fg="#2ecc71").pack()
        tk.Label(header, text=nom,
                 font=("Arial",11), bg="#1e2130", fg="#f5c842").pack()

        tk.Label(win, text="هل تريد طباعة وثائق التسجيل؟",
                 font=("Arial",10), bg=BG, fg=TEXT2).pack(pady=(14,6))

        # الأزرار الثلاثة
        bf = tk.Frame(win, bg=BG); bf.pack(pady=4, padx=20, fill="x")

        def do_tasjil():
            try:
                import طباعة_وورد as pw
                import طباعة_الشهادات as pr
                markaz = pr.load_markaz()
                path = pw.fill_tasjil_wasl(mtrsh_data, markaz, 0, 0)
                if path: pw.open_word(path)
            except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

        def do_aqd():
            try:
                import طباعة_وورد as pw
                import طباعة_الشهادات as pr
                markaz = pr.load_markaz()
                raqm = next_num(12)
                path = pw.fill_aqd_takwin(mtrsh_data, markaz, raqm)
                if path: pw.open_word(path)
            except Exception as e: messagebox.showerror("❌", f"خطأ:\n{e}")

        def do_all():
            do_tasjil()
            do_aqd()
            win.destroy()

        tk.Button(bf, text="📄  شهادة التسجيل + وصل التسديد",
                  command=do_tasjil, font=("Arial",10), bg="#2471a3", fg=WHITE,
                  relief="flat", pady=8, cursor="hand2").pack(fill="x", pady=3)

        tk.Button(bf, text="📋  عقد التكوين",
                  command=do_aqd, font=("Arial",10), bg="#8e44ad", fg=WHITE,
                  relief="flat", pady=8, cursor="hand2").pack(fill="x", pady=3)

        tk.Button(bf, text="🖨️  طباعة الكل دفعة واحدة",
                  command=do_all, font=("Arial",10,"bold"), bg="#1a6b3c", fg=WHITE,
                  relief="flat", pady=8, cursor="hand2").pack(fill="x", pady=3)

        # زر تخطي
        tk.Button(win, text="تخطي — لاحقاً", command=win.destroy,
                  font=("Arial",9), bg=BG, fg=TEXT2,
                  relief="flat", cursor="hand2").pack(pady=6)
    def _edit(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر مترشحاً"); return
        nom=self.tree.item(sel[0])["values"][0]
        rows=load_sheet("مترشحون")
        for i,r in enumerate(rows):
            if str(r[1]).strip()==str(nom).strip():
                mp=[
                    ("الاسم واللقب",1),
                    ("اللقب بالفرنسية",2),
                    ("الاسم بالفرنسية",3),
                    ("تاريخ ومكان الازدياد",4),
                    ("العنوان",5),
                    ("رقم التعريف الوطني",6),
                    ("رقم رخصة السياقة",7),
                    ("تاريخ صدور رخصة السياقة",8),
                    ("تاريخ انتهاء رخصة السياقة",9),
                    ("مكان إصدار الرخصة",10),
                    ("فئات الرخصة",11),
                    ("زمرة الدم",12),
                    ("الجنس",13),
                    ("رقم هاتف المتربص",14),
                    ("البريد الإلكتروني",15),
                    ("تخصص التكوين",16),
                    ("رقم الدورة",17),
                ]
                date_cols = {8, 9}  # تاريخ صدور وانتهاء الرخصة
                for col,cidx in mp:
                    val = str(r[cidx] or "") if cidx < len(r) else ""
                    if cidx in date_cols and val:
                        val = fmt_date(val)
                    self.mv[col].set(val)
                self._edit_row = i
                self._editing = True
                # تحميل صورة الرخصة
                nin = str(r[4] or "").strip()
                if nin: self._load_license_image(nin)
                break

    def _save_edit(self):
        """حفظ التعديل على السجل الموجود"""
        if not hasattr(self,'_editing') or not self._editing: return
        rows=load_sheet("مترشحون")
        if self._edit_row >= len(rows): return
        wb=openpyxl.load_workbook(EXCEL_FILE); ws=wb[SHEETS["مترشحون"]]
        row_num=self._edit_row+2
        # تحديث كل الأعمدة
        updates = {
            2:  "الاسم واللقب",
            3:  "اللقب بالفرنسية",
            4:  "الاسم بالفرنسية",
            5:  "تاريخ ومكان الازدياد",
            6:  "العنوان",
            7:  "رقم التعريف الوطني",
            8:  "رقم رخصة السياقة",
            9:  "تاريخ صدور رخصة السياقة",
            10: "تاريخ انتهاء رخصة السياقة",
            11: "مكان إصدار الرخصة",
            12: "فئات الرخصة",
            13: "زمرة الدم",
            14: "الجنس",
            15: "رقم هاتف المتربص",
            16: "البريد الإلكتروني",
            17: "تخصص التكوين",
            18: "نوع التكوين",
            19: "حالة النجاح",
        }
        for col, key in updates.items():
            if key in self.mv:
                ws.cell(row=row_num, column=col, value=self.mv[key].get())
        try:
            wb.save(EXCEL_FILE)
            self._editing=False
            for v in self.mv.values(): v.set("")
            self.mv["تخصص التكوين"].set(TAKHASSAS[0])
            self.refresh()
            messagebox.showinfo("✅","تم تحديث بيانات المترشح")
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح",
                "أغلق ملف Excel أولاً ثم حاول مرة أخرى")
    def _status(self,status):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر مترشحاً"); return
        nom=self.tree.item(sel[0])["values"][0]
        rows=load_sheet("مترشحون"); wb=openpyxl.load_workbook(EXCEL_FILE); ws=wb[SHEETS["مترشحون"]]
        today=datetime.now().strftime("%Y-%m-%d")
        for i,r in enumerate(rows):
            if str(r[1]).strip()==str(nom).strip():
                ws.cell(row=i+2,column=20,value=status)  # حالة النجاح
                if status=="ناجح":
                    if not (r[20] if len(r)>20 else ""):
                        ws.cell(row=i+2,column=21,value=next_num(12))
                        ws.cell(row=i+2,column=22,value=today)
                    takh_r = str(r[16] or "") if len(r)>16 else ""
                    if "البضائع" in takh_r and "الأشخاص" in takh_r:
                        if not (r[22] if len(r)>22 else ""):
                            ws.cell(row=i+2,column=23,value=next_num(14))
                            ws.cell(row=i+2,column=24,value=today)
                break
        try:
            wb.save(EXCEL_FILE)
            self.refresh()
            messagebox.showinfo("✅", f"تم تحديث حالة {nom} → {status}")
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح",
                "ملف Excel مفتوح حالياً!\n\n"
                "أغلق ملف 'بيانات_مركز_التكوين.xlsx' من Excel\n"
                "ثم حاول مرة أخرى.")
    def _whatsapp(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر مترشحاً أولاً"); return
        nom=self.tree.item(sel[0])["values"][0]
        for r in load_sheet("مترشحون"):
            if str(r[1]).strip()==str(nom).strip():
                tel=str(r[14] or "").strip().replace(" ","").replace("+","").replace("-","")
                if not tel: messagebox.showwarning("⚠️","رقم الهاتف غير مسجل"); return
                if tel.startswith("0"): tel="213"+tel[1:]
                m=get_markaz_dict(); msg="السيد/ة "+str(nom)+"،\nمن "+str(m.get("اسم مركز التكوين","المركز"))+"\nتخصص: "+str(r[16] if len(r)>16 else "")+"\nرقم الشهادة: "+str(r[20] if len(r)>20 else "في الانتظار")+"\nصالحة 5 سنوات.\nشكراً."
                webbrowser.open("https://wa.me/"+tel+"?text="+urllib.parse.quote(msg)); return
    def _email(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("warning","اختر مترشحاً أولاً"); return
        nom=self.tree.item(sel[0])["values"][0]
        for r in load_sheet("مترشحون"):
            if str(r[1]).strip()==str(nom).strip():
                email=str(r[15] or "").strip()
                if not email: messagebox.showwarning("warning","البريد الإلكتروني غير مسجل"); return
                m=get_markaz_dict()
                subj = "شهادة الكفاءة المهنية - " + str(nom)
                body = ("السيد/ة " + str(nom) + "\n\n"
                        "رقم الشهادة: " + str(r[20] if len(r)>20 else "في الانتظار") + "\n"
                        "التخصص: " + str(r[16] if len(r)>16 else "") + "\n"
                        "الشهادة صالحة 5 سنوات من تاريخ الاصدار.\n\n"
                        "مركز التكوين: " + str(m.get("اسم مركز التكوين","")) + "\n"
                        "هاتف: " + str(m.get("رقم الهاتف","")) )
                gmail = ("https://mail.google.com/mail/?view=cm&fs=1"
                         "&to=" + urllib.parse.quote(email)
                         + "&su=" + urllib.parse.quote(subj)
                         + "&body=" + urllib.parse.quote(body))
                webbrowser.open(gmail)
                return

    def _sms(self):
        """إرسال SMS عبر المودام"""
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر مترشحاً أولاً"); return
        nom=self.tree.item(sel[0])["values"][0]
        tel=""
        for r in load_sheet("مترشحون"):
            if str(r[1]).strip()==str(nom).strip():
                n=len(r)
                tel=str(r[14] or "") if n>14 else str(r[8] or "")
                break

        win=tk.Toplevel(); win.title("📱 إرسال SMS"); win.configure(bg=BG)
        win.geometry("420x400"); win.grab_set()
        tk.Label(win,text=f"📱  SMS إلى: {nom}",
                 font=("Arial",11,"bold"),bg=BG,fg=ACCENT2).pack(pady=8)

        # منفذ COM
        f1=tk.Frame(win,bg=BG); f1.pack(fill="x",padx=20,pady=3)
        tk.Label(f1,text="منفذ المودام :",bg=BG,fg=TEXT2,font=("Arial",9)).pack(side="right",padx=6)
        v_port=tk.StringVar(value="COM3")
        try:
            import قراءة_الرخصة as ql
            p=ql.detect_modem_port()
            if p: v_port.set(p)
        except Exception: pass
        tk.Entry(f1,textvariable=v_port,width=10,bg=BG3,fg=TEXT,
                 relief="flat",bd=3,justify="center").pack(side="right")

        # رقم الهاتف
        f2=tk.Frame(win,bg=BG); f2.pack(fill="x",padx=20,pady=3)
        tk.Label(f2,text="رقم الهاتف :",bg=BG,fg=TEXT2,font=("Arial",9)).pack(side="right",padx=6)
        v_tel=tk.StringVar(value=tel.strip())
        tk.Entry(f2,textvariable=v_tel,width=18,bg=BG3,fg=TEXT,
                 relief="flat",bd=3,justify="right").pack(side="right")

        # نص الرسالة
        tk.Label(win,text="نص الرسالة :",bg=BG,fg=TEXT2,
                 font=("Arial",9),anchor="e").pack(fill="x",padx=20)
        txt=tk.Text(win,height=6,bg=BG3,fg=TEXT,relief="flat",
                    bd=3,font=("Arial",10),wrap="word")
        txt.pack(fill="x",padx=20,pady=4)
        txt.insert("1.0",f"مرحباً {nom}، شهادة الكفاءة المهنية جاهزة للاستلام من مركز التكوين.")

        # عداد الأحرف
        lbl_cnt=tk.Label(win,text="",bg=BG,fg=TEXT2,font=("Arial",8)); lbl_cnt.pack()
        def upd_cnt(*a):
            n=len(txt.get("1.0","end").strip())
            lbl_cnt.config(text=f"{n} حرف / {(n//160)+1} SMS")
        txt.bind("<KeyRelease>",upd_cnt); upd_cnt()

        lbl_st=tk.Label(win,text="",bg=BG,fg=TEXT2,font=("Arial",9)); lbl_st.pack(pady=2)

        def do_send():
            phone=v_tel.get().strip(); msg=txt.get("1.0","end").strip()
            port=v_port.get().strip()
            if not phone: messagebox.showwarning("⚠️","أدخل رقم الهاتف"); return
            if not msg:   messagebox.showwarning("⚠️","أدخل نص الرسالة"); return
            lbl_st.config(text="⏳ جارٍ الإرسال...",fg="#f5c842"); win.update()
            try:
                import قراءة_الرخصة as ql
                ok,result=ql.send_sms(phone,msg,port)
                lbl_st.config(text=result,fg="#2ecc71" if ok else "#e74c3c")
                if ok: win.after(2000,win.destroy)
            except ImportError:
                messagebox.showerror("❌","pip install pyserial")
            except Exception as e:
                lbl_st.config(text=f"❌ {e}",fg="#e74c3c")

        bf=tk.Frame(win,bg=BG); bf.pack(pady=6)
        tk.Button(bf,text="📤 إرسال",command=do_send,
                  font=("Arial",10,"bold"),bg="#e67e22",fg=WHITE,
                  relief="flat",padx=16,pady=6,cursor="hand2").pack(side="right",padx=6)
        tk.Button(bf,text="❌ إلغاء",command=win.destroy,
                  font=("Arial",10),bg=BG3,fg=TEXT,
                  relief="flat",padx=10,pady=6,cursor="hand2").pack(side="right",padx=6)

    def _get_selected_mtrsh(self):
        """جلب بيانات المتربص المحدد من الجدول"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("⚠️","اختر مترشحاً أولاً"); return None
        nom = self.tree.item(sel[0])["values"][0]
        for r in load_sheet("مترشحون"):
            if str(r[1]).strip() == str(nom).strip():
                return dict(zip(COLS_MTRSH, list(r) + [""]*(len(COLS_MTRSH)-len(r))))
        return None

    def _print_word(self):
        """طباعة شهادة التسجيل + وصل التسديد بصيغة Word"""
        mtrsh = self._get_selected_mtrsh()
        if not mtrsh: return
        try:
            import طباعة_وورد as pw
        except Exception:
            messagebox.showerror("❌","ملف طباعة_وورد.py غير موجود"); return

        takh  = pw.clean_takh(str(mtrsh.get("تخصص التكوين","") or ""))
        total = pw.get_montant(takh)

        win = tk.Toplevel(); win.title("Word تسجيل + وصل"); win.configure(bg=BG)
        win.geometry("360x260"); win.grab_set()
        tk.Label(win, text="📝  شهادة التسجيل + وصل التسديد (Word)",
                 font=("Arial",11,"bold"), bg=BG, fg=ACCENT2).pack(pady=10)
        tk.Label(win, text=f"المبلغ الإجمالي : {total:,.2f} دج",
                 font=("Arial",10), bg=BG, fg=TEXT).pack(pady=2)

        f = tk.Frame(win, bg=BG); f.pack(pady=6)
        tk.Label(f, text="المسدَّد سابقاً (دج) :", bg=BG, fg=TEXT2,
                 font=("Arial",9)).grid(row=0, column=0, sticky="e", padx=8, pady=4)
        v_avant = tk.StringVar(value="0")
        tk.Entry(f, textvariable=v_avant, width=16, bg=BG3, fg=TEXT,
                 relief="flat", bd=3, justify="right").grid(row=0, column=1, pady=4)

        tk.Label(f, text="المسدَّد حالياً (دج) :", bg=BG, fg=TEXT2,
                 font=("Arial",9)).grid(row=1, column=0, sticky="e", padx=8, pady=4)
        v_now = tk.StringVar(value="0")
        tk.Entry(f, textvariable=v_now, width=16, bg=BG3, fg=TEXT,
                 relief="flat", bd=3, justify="right").grid(row=1, column=1, pady=4)

        def do_print():
            try:
                paye_now   = float(v_now.get() or 0)
                paye_avant = float(v_avant.get() or 0)
                win.destroy()
                import طباعة_الشهادات as pr
                markaz = pr.load_markaz()
                path = pw.fill_tasjil_wasl(mtrsh, markaz, paye_now, paye_avant)
                if path: pw.open_word(path)
            except Exception as e:
                messagebox.showerror("❌", f"خطأ:\n{e}")

        bf = tk.Frame(win, bg=BG); bf.pack(pady=10)
        tk.Button(bf, text="🖨️ طباعة", command=do_print,
                  font=("Arial",10,"bold"), bg="#16a085", fg=WHITE,
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right", padx=6)
        tk.Button(bf, text="❌ إلغاء", command=win.destroy,
                  font=("Arial",10), bg=BG3, fg=TEXT,
                  relief="flat", padx=10, pady=6, cursor="hand2").pack(side="right", padx=6)

    def _print_aqd(self):
        """طباعة عقد التكوين بصيغة Word"""
        mtrsh = self._get_selected_mtrsh()
        if not mtrsh: return
        try:
            import طباعة_وورد as pw
        except Exception:
            messagebox.showerror("❌","ملف طباعة_وورد.py غير موجود"); return

        win = tk.Toplevel(); win.title("عقد التكوين"); win.configure(bg=BG)
        win.geometry("320x180"); win.grab_set()
        tk.Label(win, text="📋  عقد التكوين",
                 font=("Arial",12,"bold"), bg=BG, fg=ACCENT2).pack(pady=12)
        f = tk.Frame(win, bg=BG); f.pack()
        tk.Label(f, text="رقم العقد :", bg=BG, fg=TEXT2,
                 font=("Arial",10)).grid(row=0, column=0, padx=10, pady=6)
        v_raqm = tk.StringVar(value=next_num(12))
        tk.Entry(f, textvariable=v_raqm, width=16, bg=BG3, fg=TEXT,
                 relief="flat", bd=3, justify="right").grid(row=0, column=1, pady=6)

        def do_print():
            raqm = v_raqm.get().strip() or "001"
            win.destroy()
            try:
                import طباعة_الشهادات as pr
                markaz = pr.load_markaz()
                path = pw.fill_aqd_takwin(mtrsh, markaz, raqm)
                if path: pw.open_word(path)
            except Exception as e:
                messagebox.showerror("❌", f"خطأ:\n{e}")

        bf = tk.Frame(win, bg=BG); bf.pack(pady=10)
        tk.Button(bf, text="🖨️ طباعة", command=do_print,
                  font=("Arial",10,"bold"), bg="#8e44ad", fg=WHITE,
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="right", padx=6)
        tk.Button(bf, text="❌ إلغاء", command=win.destroy,
                  font=("Arial",10), bg=BG3, fg=TEXT,
                  relief="flat", padx=10, pady=6, cursor="hand2").pack(side="right", padx=6)

    def _send_bulk(self):
        """إرسال رسالة جماعية لجميع المترشحين أو دورة محددة"""
        win = tk.Toplevel(); win.title("إرسال جماعي"); win.configure(bg=BG)
        win.geometry("520x480"); win.grab_set()

        tk.Label(win, text="📢  إرسال رسالة جماعية",
                 font=("Arial",13,"bold"), bg=BG, fg=ACCENT2).pack(pady=10)

        # فلتر الدورة
        f1 = tk.Frame(win, bg=BG2, pady=8, padx=15); f1.pack(fill="x", padx=15)
        tk.Label(f1, text="الدورة:", bg=BG2, fg=TEXT2,
                 font=("Arial",10)).pack(side="right", padx=8)
        dawra_var = tk.StringVar(value="الكل")
        dawras = ["الكل"] + list({fmt_raqm(str(r[17])) for r in load_sheet("مترشحون") if len(r)>17 and r[17]})
        ttk.Combobox(f1, textvariable=dawra_var, values=dawras,
                     width=14, state="readonly").pack(side="right", padx=4)

        tk.Label(f1, text="الحالة:", bg=BG2, fg=TEXT2,
                 font=("Arial",10)).pack(side="right", padx=8)
        status_var = tk.StringVar(value="الكل")
        ttk.Combobox(f1, textvariable=status_var,
                     values=["الكل","ناجح","راسب","غائب","مؤجل","متخلي","معيد"], width=10,
                     state="readonly").pack(side="right", padx=4)

        # نص الرسالة
        tk.Label(win, text="نص الرسالة:", font=("Arial",9,"bold"),
                 bg=BG, fg=TEXT2).pack(anchor="e", padx=20, pady=(8,2))
        m = get_markaz_dict()
        default_msg = ("من: " + str(m.get("اسم مركز التكوين","المركز")) + "\n"
                      "نعلمكم بأن نتائج الدورة التكوينية أصبحت متاحة.\n"
                      "للاستفسار: " + str(m.get("رقم الهاتف","")) )
        txt = tk.Text(win, height=6, width=56, bg=BG2, fg=TEXT,
                      font=("Arial",9), relief="flat", bd=4, wrap="word")
        txt.pack(padx=20); txt.insert("1.0", default_msg)

        # قناة الإرسال
        tk.Label(win, text="قناة الإرسال:", font=("Arial",9,"bold"),
                 bg=BG, fg=TEXT2).pack(pady=(8,2))
        channel_var = tk.StringVar(value="whatsapp")
        cf = tk.Frame(win, bg=BG); cf.pack()
        for txt_btn, val, color in [
            ("💬 واتساب","whatsapp","#25D366"),
            ("✉️ Gmail","gmail","#c0392b"),
            ("📱 SMS موبيليس","sms","#e67e22"),
        ]:
            tk.Radiobutton(cf, text=txt_btn, variable=channel_var, value=val,
                          bg=BG, fg=color, selectcolor=BG2,
                          font=("Arial",10,"bold"),
                          activebackground=BG).pack(side="left", padx=10)

        # معلومات
        lbl_info = tk.Label(win, text="", font=("Arial",8), bg=BG, fg=TEXT2)
        lbl_info.pack()

        def preview():
            rows = load_sheet("مترشحون")
            fd = dawra_var.get(); fs = status_var.get()
            filtered = [r for r in rows
                       if (fd=="الكل" or fmt_raqm(str(r[17] if len(r)>17 else "")).strip()==str(fd).strip())
                       and (fs=="الكل" or (r[19] if len(r)>19 else "")==fs)]
            ch = channel_var.get()
            has_contact = sum(1 for r in filtered
                            if (ch=="gmail" and len(r)>15 and r[15]) or
                               (ch in ("whatsapp","sms") and len(r)>14 and r[14]))
            lbl_info.config(
                text=f"سيتم الإرسال لـ {has_contact} مترشح من أصل {len(filtered)}")

        channel_var.trace_add("write", lambda *a: preview())
        dawra_var.trace_add("write", lambda *a: preview())
        status_var.trace_add("write", lambda *a: preview())
        preview()

        def do_send():
            msg_text = txt.get("1.0","end").strip()
            rows = load_sheet("مترشحون")
            fd = dawra_var.get(); fs = status_var.get()
            ch = channel_var.get()
            filtered = [r for r in rows
                       if (fd=="الكل" or fmt_raqm(str(r[17] if len(r)>17 else "")).strip()==str(fd).strip())
                       and (fs=="الكل" or (r[19] if len(r)>19 else "")==fs)]
            sent = 0
            import urllib.parse
            for r in filtered:
                nom = str(r[1] or "")
                n_r = len(r)
                personal_msg = msg_text.replace("{الاسم}", nom).replace("{الشهادة}", str(r[20] if n_r>19 else ""))
                if ch == "whatsapp":
                    tel = str(r[14] if n_r>14 else "").strip().replace(" ","").replace("+","").replace("-","")
                    if not tel: continue
                    if tel.startswith("0"): tel = "213" + tel[1:]
                    webbrowser.open("https://wa.me/"+tel+"?text="+urllib.parse.quote(personal_msg))
                    sent += 1
                elif ch == "gmail":
                    email = str(r[15] if n_r>15 else "").strip()
                    if not email: continue
                    subj = "إشعار من مركز التكوين"
                    url = ("https://mail.google.com/mail/?view=cm&fs=1"
                           "&to=" + urllib.parse.quote(email)
                           + "&su=" + urllib.parse.quote(subj)
                           + "&body=" + urllib.parse.quote(personal_msg))
                    webbrowser.open(url)
                    sent += 1
                elif ch == "sms":
                    n_r2 = len(r)
                    tel = str(r[14] if n_r2 > 14 else "").strip()
                    if not tel: continue
                    try:
                        import قراءة_الرخصة as ql
                        ok, result = ql.send_sms(tel, personal_msg)
                        if ok: sent += 1
                        else:
                            # احتياطي: نسخ للحافظة
                            win.clipboard_clear()
                            win.clipboard_append(personal_msg)
                    except Exception:
                        win.clipboard_clear()
                        win.clipboard_append(personal_msg)
                        sent += 1

            win.destroy()
            messagebox.showinfo("✅ تم",
                "تم فتح " + str(sent) + " رسالة\n"
                "ملاحظة: يمكنك استخدام {الاسم} و{الشهادة} في النص")

        bf = tk.Frame(win, bg=BG); bf.pack(pady=10)
        tk.Button(bf, text="📢 إرسال للجميع", command=do_send,
                  font=("Arial",11,"bold"), bg=SUCCESS, fg=WHITE,
                  relief="flat", padx=15, pady=8, cursor="hand2").pack(side="right",padx=5)
        tk.Button(bf, text="❌ إلغاء", command=win.destroy,
                  font=("Arial",10), bg=BG3, fg=TEXT,
                  relief="flat", padx=10, pady=8, cursor="hand2").pack(side="right",padx=5)

    def _import(self):
        path=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")],title="استيراد Excel")
        if not path: return
        try:
            wb_s=openpyxl.load_workbook(path); ws_s=wb_s.active
            hdrs=[str(ws_s.cell(1,c).value or "").strip() for c in range(1,ws_s.max_column+1)]
            MAP={"الاسم":"الاسم واللقب","الاسم واللقب":"الاسم واللقب","تاريخ الميلاد":"تاريخ ومكان الازدياد","تاريخ ومكان الازدياد":"تاريخ ومكان الازدياد","العنوان":"العنوان","رقم التعريف":"رقم التعريف الوطني","رقم التعريف الوطني":"رقم التعريف الوطني","رقم الرخصة":"رقم رخصة السياقة","رقم رخصة السياقة":"رقم رخصة السياقة","تاريخ الرخصة":"تاريخ صدور رخصة السياقة","تاريخ صدور رخصة السياقة":"تاريخ صدور رخصة السياقة","التخصص":"تخصص التكوين","تخصص التكوين":"تخصص التكوين","رقم الدورة":"رقم الدورة","الهاتف":"رقم هاتف المتربص","رقم هاتف المتربص":"رقم هاتف المتربص","البريد":"البريد الإلكتروني","البريد الإلكتروني":"البريد الإلكتروني"}
            added=skipped=0
            for row in ws_s.iter_rows(min_row=2,values_only=True):
                if not any(row): continue
                d=dict(zip(hdrs,row))
                def g(f): return next((str(v) if v else "" for k,v in d.items() if MAP.get(k)==f or k==f),"")
                nom=g("الاسم واللقب")
                if not nom: skipped+=1; continue
                save_row("مترشحون",[datetime.now().strftime("%Y-%m-%d"),nom,g("تاريخ ومكان الازدياد"),g("العنوان"),g("رقم التعريف الوطني"),g("رقم رخصة السياقة"),g("تاريخ صدور رخصة السياقة"),g("تخصص التكوين") or TAKHASSAS[0],g("رقم هاتف المتربص"),g("البريد الإلكتروني"),g("رقم الدورة"),"","","","",""]); added+=1
            self.refresh(); messagebox.showinfo("✅ الاستيراد",f"تم استيراد {added} مترشح\\nتجاهل {skipped} فارغ")
        except Exception as e: messagebox.showerror("❌ خطأ","فشل الاستيراد:\\n"+str(e))
    def _export_tpl(self):
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile="قالب_استيراد_المترشحين.xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="المترشحون"; ws.sheet_view.rightToLeft=True
        cols=["الاسم واللقب","تاريخ ومكان الازدياد","العنوان","رقم التعريف الوطني","رقم رخصة السياقة","تاريخ صدور رخصة السياقة","تخصص التكوين","رقم الدورة","رقم هاتف المتربص","البريد الإلكتروني"]
        hf,hfill,halign,_=_hstyle()
        for i,h in enumerate(cols,1):
            c=ws.cell(row=1,column=i,value=h); c.font=hf; c.fill=hfill; c.alignment=halign; ws.column_dimensions[get_column_letter(i)].width=24
        ws.append(["الاسم واللقب","تاريخ-مكان الازدياد","العنوان","رقم التعريف الوطني","رقم رخصة السياقة","تاريخ صدور رخصة السياقة","تخصص التكوين","رقم الدورة","رقم هاتف المتربص","البريد الإلكتروني"])
        # صف مثال باللون الرمادي
        from openpyxl.styles import Font as _F, PatternFill as _PF
        ex_row = ["مثال: محمد أحمد","1990/03/15-البيض","البيض","198003150001","A00000001","2015/06/01","الأشخاص","01","0661234567",""]
        ws.append(ex_row)
        for cell in ws[3]:
            cell.font = _F(color="AAAAAA", italic=True)
            cell.fill = _PF("solid", start_color="F5F5F5", end_color="F5F5F5")
        wb.save(path); messagebox.showinfo("✅","تم تصدير القالب")

class ShahadaPage(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=BG); self.app=app; self._build()
    def _build(self):
        tk.Label(self,text="📄  طباعة الشهادات",font=_efont(14,bold=True),bg=BG,fg=ACCENT2).pack(pady=(8,2))
        ff=tk.Frame(self,bg=BG2,pady=7); ff.pack(fill="x",padx=20)
        tk.Label(ff,text="رقم الدورة:",bg=BG2,fg=TEXT2).pack(side="right",padx=6)
        self.fd=tk.StringVar(value="الكل"); self.cb=ttk.Combobox(ff,textvariable=self.fd,width=14,state="normal")
        self.cb.pack(side="right",padx=4); self.cb.bind("<<ComboboxSelected>>",lambda e:self.refresh())
        tk.Label(ff,text="التخصص:",bg=BG2,fg=TEXT2).pack(side="right",padx=6)
        self.ft=tk.StringVar(value="الكل"); ttk.Combobox(ff,textvariable=self.ft,values=["الكل"]+TAKHASSAS,width=18,state="readonly").pack(side="right",padx=4)
        sbtn(ff,"‏🔍 فلترة",self.refresh,ACCENT,pady=5).pack(side="right",padx=6)
        tk.Label(ff,text="أو أدخل رقم الدورة:",bg=BG2,fg=TEXT2).pack(side="left",padx=10)
        self.di=tk.StringVar(); tk.Entry(ff,textvariable=self.di,width=12,bg=BG3,fg=TEXT,relief="flat",bd=3,insertbackground=TEXT,justify="center").pack(side="left",padx=4)
        sbtn(ff,"بحث",self._search_input,BG3,pady=5).pack(side="left",padx=4)
        tbl=tk.Frame(self,bg=BG); tbl.pack(fill="both",expand=True,padx=20,pady=4)
        cols=["الاسم واللقب","التخصص","رقم الدورة","رقم الشهادة","تاريخ الشهادة","صالحة حتى"]
        self.tree,sb=mk_tree(tbl,cols,14)
        for col,w in zip(cols,[220,120,90,130,120,120]):
            self.tree.heading(col,text=col,anchor="center")
            self.tree.column(col,width=w,minwidth=60,anchor="center",stretch=True)
        self.tree.pack(side="right",fill="both",expand=True); sb.pack(side="left",fill="y")
        bf=tk.Frame(self,bg=BG,pady=6); bf.pack(fill="x",padx=20)
        sbtn(bf,"‏🖨️ طباعة المحدد",self._print_one,ACCENT).pack(side="right",padx=5)
        sbtn(bf,"‏🖨️ طباعة كل الدورة",self._print_all,"#8e44ad").pack(side="right",padx=5)
        sbtn(bf,"‏📤 تصدير Excel",self._export,SUCCESS).pack(side="left",padx=5)
        # ── أزرار تغيير الحالة ────────────────────────────────
        bf2=tk.Frame(self,bg=BG,pady=3); bf2.pack(fill="x",padx=20)
        tk.Label(bf2,text="تغيير حالة المترشح المحدد:",bg=BG,fg=TEXT2,font=("Arial",9)).pack(side="right",padx=6)
        for lbl,clr,st in [
            ("✅ ناجح",  SUCCESS,  "ناجح"),
            ("❌ راسب",  DANGER,   "راسب"),
            ("🚫 غائب",  "#7f8c8d","غائب"),
            ("⏸ مؤجل",  "#e67e22","مؤجل"),
            ("🚪 متخلي", "#c0392b","متخلي"),
            ("🔄 معيد",  "#2980b9","معيد"),
        ]:
            sbtn(bf2,lbl,lambda s=st:self._status(s),clr,pady=4,padx=8).pack(side="right",padx=3)
        self.sd=[]
    def _search_input(self):
        v=self.di.get().strip()
        if v: self.fd.set(v)
        self.refresh()
    def refresh(self):
        for r in self.tree.get_children(): self.tree.delete(r)
        dawras=set(); self.sd=[]
        fd=self.fd.get(); ft=self.ft.get()
        for r in load_sheet("مترشحون"):
            n = len(r)
            # دعم الهيكل القديم (16 عمود) والجديد (25 عمود)
            if n <= 16:
                # هيكل قديم: تخصص=7، دورة=10، حالة=11، شهادة=12، تاريخ=13
                status = str(r[11] or "").strip() if n > 11 else ""
                takh   = str(r[7]  or "") if n > 7  else ""
                dawra  = str(r[10] or "") if n > 10 else ""
                num_sh = str(r[12] or "") if n > 12 else ""
                date_sh= str(r[13] or "") if n > 13 else ""
            else:
                # هيكل جديد
                status = str(r[19] or "").strip() if n > 18 else ""
                takh   = str(r[16] or "") if n > 16 else ""
                dawra  = str(r[17] or "") if n > 17 else ""
                num_sh = str(r[20] or "") if n > 19 else ""
                date_sh= str(r[21] or "") if n > 20 else ""

            if status != "ناجح": continue
            if dawra: dawras.add(dawra)
            if fd!="الكل" and dawra.strip()!=str(fd).strip(): continue
            if ft!="الكل" and takh!=str(ft): continue
            exp = calc_exp(date_sh) if date_sh else "—"
            self.tree.insert("","end",values=(
                r[1],
                takh,
                fmt_raqm(dawra),
                num_sh or "—",
                fmt_date(date_sh) or "—",
                fmt_date(exp) or "—"
            ))
            self.sd.append(r)
            # مزدوج
            if "البضائع" in takh and "الأشخاص" in takh:
                num2  = str(r[22] or "") if n > 21 else ""
                date2 = str(r[23] or "") if n > 22 else ""
                exp2  = fmt_date(calc_exp(date2)) if date2 else "—"
                if num2:
                    self.tree.insert("","end",values=(str(r[1])+" (2)","نقل البضائع",fmt_raqm(dawra),num2,fmt_date(date2) or "—",exp2))
        self.cb["values"]=["الكل"]+sorted(dawras)
    def _print_one(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("⚠️","اختر شهادة أولاً"); return
        idx=self.tree.index(sel[0])
        if idx<len(self.sd): self._print_pdf([self.sd[idx]])
    def _print_all(self):
        if not self.sd: messagebox.showwarning("⚠️","لا توجد شهادات"); return
        self._print_pdf(self.sd)
    def _print_pdf(self, lst):
        try:
            import طباعة_الشهادات as pr
        except ImportError:
            messagebox.showerror("❌","ملف طباعة_الشهادات.py غير موجود"); return
        if not lst:
            messagebox.showwarning("⚠️","لا توجد شهادات للطباعة"); return
        try:
            pr._register_font()
        except Exception as e:
            messagebox.showerror("❌","خطأ في الخط: "+str(e)); return
        _out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "شهادات_PDF")
        os.makedirs(_out_dir, exist_ok=True)
        m = get_markaz_dict()
        from reportlab.pdfgen import canvas as rlc
        from reportlab.lib.pagesizes import A5
        W2, H2 = A5
        out = os.path.join(_out_dir, "شهادات_"+datetime.now().strftime("%Y%m%d_%H%M%S")+".pdf")
        cv = rlc.Canvas(out, pagesize=(W2, H2))
        cv.setPageSize((W2, H2))
        page = 0
        # تحميل معلومات الدورة
        dawra_info = {}
        try:
            fd_val = self.fd.get() if hasattr(self,"fd") else "الكل"
            dawra_info = pr.load_dawra_info(fd_val)
        except Exception:
            pass
        for r in lst:
            # r قد يكون list أو dict
            def g(key, idx_num):
                if isinstance(r, dict): return r.get(key,"") or ""
                return (r[idx_num] if len(r)>idx_num else "") or ""
            takh     = str(g("تخصص التكوين", 16))
            takh     = takh.replace("نقل الأشخاص و البضائع","الأشخاص و البضائع").replace("نقل ","")
            sign_raw = str(g("تاريخ الشهادة", 21) or datetime.now().strftime("%Y-%m-%d"))
            num_sh   = str(g("رقم الشهادة", 20) or "")
            shs = [(takh if takh != "الأشخاص و البضائع" else "الأشخاص", num_sh, sign_raw)]
            num2 = str(g("رقم الشهادة 2", 22) or "")
            if takh == "الأشخاص و البضائع" and num2:
                shs.append(("البضائع", num2, str(g("تاريخ الشهادة 2",23) or sign_raw)))
            for takh_sh, ns, sd in shs:
                if page > 0: cv.showPage()
                page += 1
                data = {
                    "wilaya":    pr._clean_wilaya(m.get("مديرية النقل لولاية","")),
                    "docnum":    ns,
                    "director":  pr._clean_wilaya(m.get("مديرية النقل لولاية","")),
                    "name":      str(g("الاسم واللقب",1)),
                    "dob":       str(g("تاريخ ومكان الازدياد",4)),
                    "address":   str(g("العنوان",5)),
                    "vehicles":  takh_sh,
                    "date_from": pr.format_date(sd),
                    "date_to":   pr.calc_expiry(sd),
                    "city":      pr._clean_wilaya(m.get("مديرية النقل لولاية","")),
                    "sign_date": pr.format_date(sd),
                }
                for fid,(top,right,fsize) in pr.FIELDS.items():
                    pr.draw_field(cv, data.get(fid,""), top, right, fsize,
                                  center=(fid=="vehicles"))
                # QR Code
                try:
                    qr_data = dict(data)
                    qr_data["centre"]     = m.get("اسم مركز التكوين","")
                    qr_data["raqm_dawra"] = str(g("رقم الدورة",17))
                    qr_data["date_ouv"]   = dawra_info.get("date_ouv","")
                    qr_data["date_clo"]   = dawra_info.get("date_clo","")
                    pr.draw_qr(cv, qr_data)
                except Exception:
                    pass
        cv.save()
        if messagebox.askyesno("✅ تم", str(page)+" شهادة\n"+out+"\nفتح الملف؟"):
            pr.open_pdf(out)
        # ── رسائل تلقائية بعد الطباعة ────────────────────────
        self._send_congrats(lst, m)

    def _send_congrats(self, lst, markaz):
        """إرسال رسائل تهنئة تلقائية للمترشحين عبر واتساب + SMS + إيميل"""
        if not lst: return
        if not messagebox.askyesno("📨 رسائل تهنئة",
            f"هل تريد إرسال رسائل تهنئة لـ {len(lst)} مترشح؟\n"
            "واتساب + SMS + بريد إلكتروني"):
            return

        centre = str(markaz.get("اسم مركز التكوين","المركز"))
        errors = []

        for r in lst:
            def g(key, idx_num):
                if isinstance(r, dict): return str(r.get(key,"") or "")
                return str(r[idx_num] if len(r)>idx_num else "") or ""

            nom      = g("الاسم واللقب", 1)
            takh     = g("تخصص التكوين", 16).replace("نقل ","")
            num_sh   = g("رقم الشهادة", 20)
            date_sh  = g("تاريخ الشهادة", 21)
            tel      = g("رقم هاتف المتربص", 14).strip()
            email    = g("البريد الإلكتروني", 15).strip()

            # حساب تاريخ الصلاحية
            try:
                import طباعة_الشهادات as pr
                date_to = pr.calc_expiry(date_sh)
                date_from = pr.format_date(date_sh)
            except Exception:
                date_from = date_sh; date_to = ""

            # نص الرسالة
            msg = (
                f"🎓 تهنئة بالنجاح\n\n"
                f"السيد/السيدة: {nom}\n\n"
                f"يسعدنا إبلاغكم بأنكم اجتزتم بنجاح التكوين المهني\n"
                f"في تخصص: {takh}\n"
                f"بمركز: {centre}\n\n"
                f"رقم الشهادة: {num_sh}\n"
                f"صالحة من: {date_from}\n"
                f"إلى: {date_to}\n\n"
                f"مبروك وبالتوفيق 🎉"
            )

            # ── واتساب ──────────────────────────────────────
            if tel:
                t = tel.replace("+","").replace(" ","").replace("-","")
                if t.startswith("0"): t = "213" + t[1:]
                wa_url = "https://wa.me/" + t + "?text=" + urllib.parse.quote(msg)
                webbrowser.open(wa_url)

            # ── بريد إلكتروني ────────────────────────────────
            if email:
                subj = f"تهنئة بالحصول على شهادة الكفاءة المهنية — {nom}"
                gmail = ("https://mail.google.com/mail/?view=cm&fs=1"
                         "&to=" + urllib.parse.quote(email)
                         + "&su=" + urllib.parse.quote(subj)
                         + "&body=" + urllib.parse.quote(msg))
                webbrowser.open(gmail)

            # ── SMS ──────────────────────────────────────────
            if tel:
                sms_msg = (f"تهنئة {nom} — شهادة كفاءة {takh} — رقم: {num_sh} — صالحة حتى: {date_to} — {centre}")
                # نسخ آخر رسالة SMS للحافظة
                try:
                    import pyperclip
                    pyperclip.copy(sms_msg)
                except Exception:
                    pass

        messagebox.showinfo("✅ الرسائل",
            f"تم فتح واتساب وإيميل لـ {len(lst)} مترشح.\n"
            "للـ SMS: افتح موبيليس والصق الرسالة من الحافظة.")


    def _export(self):
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="الناجحون"; ws.sheet_view.rightToLeft=True
        hs=["الاسم واللقب","تاريخ ومكان الازدياد","العنوان","التخصص","رقم الدورة","رقم الشهادة","تاريخ الشهادة","صالحة حتى"]
        hf,hfill,halign,_=_hstyle()
        for i,h in enumerate(hs,1):
            c=ws.cell(row=1,column=i,value=h); c.font=hf; c.fill=hfill; c.alignment=halign; ws.column_dimensions[get_column_letter(i)].width=22
        for r in self.sd:
            td=str(r[21] or "") if len(r)>21 else ""
            ws.append([r[1],r[4] if len(r)>4 else "",r[5] if len(r)>5 else "",r[16] if len(r)>16 else "",r[17] if len(r)>17 else "",r[20] if len(r)>20 else "",fmt_date(td),calc_exp(td)])
        wb.save(path); messagebox.showinfo("✅","تم تصدير "+str(len(self.sd))+" شهادة")

class StatsPage(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=BG); self.app=app; self._build()
    def _build(self):
        tk.Label(self,text="📊  الإحصائيات والتقارير الشهرية",font=_efont(14,bold=True),bg=BG,fg=ACCENT2).pack(pady=(8,4))
        top=tk.Frame(self,bg=BG2,pady=7); top.pack(fill="x",padx=20)
        tk.Label(top,text="الشهر:",bg=BG2,fg=TEXT,font=("Arial",10,"bold")).pack(side="right",padx=8)
        self.mv=tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        tk.Entry(top,textvariable=self.mv,width=10,bg=BG3,fg=TEXT,relief="flat",bd=3,insertbackground=TEXT,justify="right",font=("Arial",11)).pack(side="right",padx=4)
        tk.Label(top,text="YYYY-MM",bg=BG2,fg=TEXT2,font=("Arial",8)).pack(side="right")
        sbtn(top,"‏📊 عرض",self.refresh,ACCENT,pady=5).pack(side="right",padx=8)
        sbtn(top,"‏📤 تصدير",self._export,SUCCESS,pady=5).pack(side="left",padx=8)
        self.cf=tk.Frame(self,bg=BG); self.cf.pack(fill="x",padx=20,pady=8)
        self.df=tk.Frame(self,bg=BG); self.df.pack(fill="both",expand=True,padx=20,pady=4)
    def refresh(self):
        for w in self.cf.winfo_children(): w.destroy()
        for w in self.df.winfo_children(): w.destroy()
        month=self.mv.get().strip(); rows=load_sheet("مترشحون")
        flt=[r for r in rows if str(r[0] or "").startswith(month)]
        # r[19]=حالة النجاح  r[16]=تخصص التكوين  r[17]=رقم الدورة
        total=len(flt)
        najihin =sum(1 for r in flt if len(r)>19 and r[19]=="ناجح")
        rasibin =sum(1 for r in flt if len(r)>19 and r[19]=="راسب")
        ashkhas =sum(1 for r in flt if len(r)>16 and r[16] and "أشخاص" in str(r[16]))
        bada2i  =sum(1 for r in flt if len(r)>16 and r[16] and "بضائع" in str(r[16]))
        khatira =sum(1 for r in flt if len(r)>16 and r[16] and "خطرة"  in str(r[16]))
        muzdawaj=sum(1 for r in flt if len(r)>16 and r[16]=="الأشخاص و البضائع")
        nisbah=round(najihin/total*100,1) if total else 0
        cards=[("📋 المترشحون",total,ACCENT),("✅ الناجحون",najihin,SUCCESS),("❌ الراسبون",rasibin,DANGER),(f"🎯 {nisbah}%","نسبة النجاح",WARNING),("🚌 الأشخاص",ashkhas,"#3498db"),("📦 البضائع",bada2i,"#9b59b6"),("☣️ الخطرة",khatira,"#e67e22"),("🔀 مزدوج",muzdawaj,"#1abc9c")]
        for i,(label,val,color) in enumerate(cards):
            card=tk.Frame(self.cf,bg=BG2,padx=12,pady=10); card.grid(row=0,column=i,padx=4,pady=2,sticky="ew"); self.cf.columnconfigure(i,weight=1)
            tk.Label(card,text=str(val),font=("Arial",20,"bold"),bg=BG2,fg=color).pack()
            tk.Label(card,text=label,font=("Arial",7),bg=BG2,fg=TEXT2).pack()
            tk.Frame(card,bg=color,height=3).pack(fill="x",pady=(5,0))
        tk.Label(self.df,text="تفصيل حسب الدورة:",font=("Arial",10,"bold"),bg=BG,fg=TEXT).pack(anchor="e",pady=4)
        ds={}
        for r in flt:
            n=len(r)
            d=fmt_raqm(str(r[17] or "غير محدد")) if n>17 else "غير محدد"
            if d not in ds: ds[d]={"إجمالي":0,"ناجح":0,"راسب":0}
            ds[d]["إجمالي"]+=1
            hal=str(r[19] or "") if n>19 else ""
            if hal=="ناجح":  ds[d]["ناجح"]+=1
            elif hal=="راسب": ds[d]["راسب"]+=1
        cols=["رقم الدورة","الإجمالي","الناجحون","الراسبون","نسبة النجاح"]
        tree,sb=mk_tree(self.df,cols,8)
        for col in cols: tree.heading(col,text=col); tree.column(col,width=150,anchor="center")
        for dawra,s in ds.items():
            ns=round(s["ناجح"]/s["إجمالي"]*100,1) if s["إجمالي"] else 0
            tree.insert("","end",values=(dawra,s["إجمالي"],s["ناجح"],s["راسب"],str(ns)+"%"))
        tree.pack(side="right",fill="both",expand=True); sb.pack(side="left",fill="y")
    def _export(self):
        month=self.mv.get().strip()
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",initialfile="تقرير_"+month+".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=load_sheet("مترشحون"); flt=[r for r in rows if str(r[0] or "").startswith(month)]
        m=get_markaz(); wb=openpyxl.Workbook()
        ws1=wb.active; ws1.title="ملخص"; ws1.sheet_view.rightToLeft=True
        ws1["A1"]="التقرير الشهري — "+month; ws1["A1"].font=Font(bold=True,size=14,color="1F3864")
        ws1["A2"]="مركز التكوين: "+(str(m[0]) if m else ""); ws1["A3"]="مديرية النقل: "+(str(m[4]) if len(m)>4 else "")
        ws1["A4"]="التاريخ: "+datetime.now().strftime("%Y-%m-%d %H:%M")
        # r[19]=حالة النجاح  r[16]=تخصص  r[17]=دورة  r[20]=شهادة  r[21]=تاريخ الشهادة
        total=len(flt)
        najihin=sum(1 for r in flt if len(r)>19 and r[19]=="ناجح")
        rasibin=sum(1 for r in flt if len(r)>19 and r[19]=="راسب")
        for i,(label,val) in enumerate([("الإجمالي",total),("الناجحون",najihin),("الراسبون",rasibin),("نسبة النجاح",str(round(najihin/total*100,1) if total else 0)+"%")],6):
            ws1.cell(row=i,column=1,value=label); ws1.cell(row=i,column=2,value=val)
        ws2=wb.create_sheet("تفاصيل"); ws2.sheet_view.rightToLeft=True
        hs=["الاسم","تاريخ ومكان الازدياد","العنوان","التخصص","الدورة","الحالة","الشهادة","التاريخ"]
        hf,hfill,halign,_=_hstyle()
        for i,h in enumerate(hs,1):
            c=ws2.cell(row=1,column=i,value=h); c.font=hf; c.fill=hfill; c.alignment=halign; ws2.column_dimensions[get_column_letter(i)].width=20
        for r in flt:
            n=len(r)
            ws2.append([r[1] if n>1 else "",r[4] if n>4 else "",r[5] if n>5 else "",
                        r[16] if n>16 else "",fmt_raqm(str(r[17] or "")) if n>17 else "",
                        r[19] if n>19 else "",r[20] if n>20 else "",fmt_date(str(r[21] or "")) if n>21 else ""])
        wb.save(path); messagebox.showinfo("✅","تم تصدير التقرير:\n"+path)


class DawraGestionPage(tk.Frame):
    """صفحة متكاملة لتسيير دورة واحدة: عرض المترشحين، تغيير الحالات، إحصائيات فورية"""

    STATUTS = ["ناجح","راسب","غائب","مؤجل","متخلي","معيد",""]
    COLORS  = {
        "ناجح":  "#2ecc71",
        "راسب":  "#e74c3c",
        "غائب":  "#95a5a6",
        "مؤجل":  "#f39c12",
        "متخلي": "#e67e22",
        "معيد":  "#3498db",
        "":      TEXT2,
    }
    TAG_BG  = {
        "ناجح":  "#1a3a2a",
        "راسب":  "#3a1a1a",
        "غائب":  "#2a2a2a",
        "مؤجل":  "#2a2000",
        "متخلي": "#3a1a00",
        "معيد":  "#001a3a",
        "":      BG2,
    }

    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app      = app
        self.raqm_var = tk.StringVar()
        self._build()

    # ─── بناء الواجهة ────────────────────────────────────────────────────────
    def _build(self):
        # ── شريط العنوان + اختيار الدورة ──────────────────────────────────
        hdr = tk.Frame(self, bg=BG2, pady=6)
        hdr.pack(fill="x", padx=0, pady=(0,2))

        tk.Label(hdr, text="🗂  تسيير الدورة", font=_efont(13,bold=True),
                 bg=BG2, fg=ACCENT2).pack(side="right", padx=20)
        sbtn(hdr, "‏🔄", self.refresh, BG3, padx=8, pady=3).pack(side="left", padx=6)
        tk.Label(hdr, text="رقم الدورة:", bg=BG2, fg=TEXT2,
                 font=("Arial",10)).pack(side="right", padx=(20,4))
        self.cb_dawra = ttk.Combobox(hdr, textvariable=self.raqm_var,
                                     width=14, state="readonly", font=("Arial",10))
        self.cb_dawra.pack(side="right", padx=4)
        self.cb_dawra.bind("<<ComboboxSelected>>", lambda e: self._load_dawra())

        # ── بطاقات الإحصائيات (مدمجة صغيرة) ─────────────────────────────
        self.stats_frame = tk.Frame(self, bg=BG)
        self.stats_frame.pack(fill="x", padx=16, pady=(2,2))
        self._stat_labels = {}
        stats_defs = [
            ("الإجمالي", ACCENT,    "total"),
            ("ناجح",     "#2ecc71", "ناجح"),
            ("راسب",     "#e74c3c", "راسب"),
            ("غائب",     "#95a5a6", "غائب"),
            ("مؤجل",     "#f39c12", "مؤجل"),
            ("متخلي",    "#e67e22", "متخلي"),
            ("معيد",     "#3498db", "معيد"),
            ("بدون حالة",TEXT2,     ""),
        ]
        for i, (label, color, key) in enumerate(stats_defs):
            card = tk.Frame(self.stats_frame, bg=BG3, padx=8, pady=3)
            card.grid(row=0, column=i, padx=3, sticky="ew")
            self.stats_frame.columnconfigure(i, weight=1)
            tk.Label(card, text=label, bg=BG3, fg=TEXT2, font=("Arial",7)).pack()
            lbl = tk.Label(card, text="—", bg=BG3, fg=color, font=("Arial",14,"bold"))
            lbl.pack()
            self._stat_labels[key] = lbl

        # ── شريط أزرار الحالة + المحاضر ──────────────────────────────────
        abar = tk.Frame(self, bg=BG2, pady=4)
        abar.pack(fill="x", padx=16, pady=(2,2))

        # أزرار الحالة
        btn_defs = [
            ("ناجح",  "✅", "#2ecc71", "ناجح"),
            ("راسب",  "❌", "#e74c3c", "راسب"),
            ("غائب",  "🚫", "#7f8c8d", "غائب"),
            ("مؤجل",  "⏸",  "#f39c12", "مؤجل"),
            ("متخلي", "🚪", "#e67e22", "متخلي"),
            ("معيد",  "🔄", "#3498db", "معيد"),
            ("مسح",   "↩",  BG3,       ""),
        ]
        tk.Label(abar, text="الحالة:", bg=BG2, fg=TEXT2,
                 font=("Arial",8)).pack(side="right", padx=(10,2))
        for label, ico, clr, st in btn_defs:
            f = tk.Frame(abar, bg=clr, cursor="hand2")
            f.pack(side="right", padx=2, pady=1)
            tk.Label(f, text=ico,   bg=clr, fg=WHITE,
                     font=_efont(10), padx=1).pack(side="right")
            tk.Label(f, text=label, bg=clr, fg=WHITE,
                     font=("Arial", 8, "bold"), padx=3, pady=3).pack(side="right")
            f.bind("<Button-1>", lambda e, s=st: self._set_status(s))
            for w in f.winfo_children():
                w.bind("<Button-1>", lambda e, s=st: self._set_status(s))

        # أزرار المحاضر
        tk.Frame(abar, bg=TEXT2, width=1).pack(side="right", fill="y", padx=6)
        sbtn(abar, "‏📋 مداولة",   self._mahdar_mudalawa, "#117a65", pady=3, padx=6).pack(side="right", padx=2)
        sbtn(abar, "‏📄 فتح",      self._mahdar_fath,     "#2471a3", pady=3, padx=6).pack(side="right", padx=2)
        sbtn(abar, "‏📊 الناجحين", self._liste_najihin,   "#1a5276", pady=3, padx=6).pack(side="right", padx=2)

        # ── شريط الفلترة + البحث ──────────────────────────────────────────
        fbar = tk.Frame(self, bg=BG2, pady=4)
        fbar.pack(fill="x", padx=16, pady=(0,2))

        tk.Label(fbar, text="🔍", bg=BG2, fg=TEXT2,
                 font=("Arial",9)).pack(side="right", padx=(10,2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *a: self._filter())
        tk.Entry(fbar, textvariable=self._search_var, width=18,
                 bg=BG3, fg=TEXT, relief="flat", bd=3,
                 insertbackground=TEXT, justify="right",
                 font=("Arial",9)).pack(side="right", padx=4)

        tk.Label(fbar, text="الحالة:", bg=BG2, fg=TEXT2,
                 font=("Arial",9)).pack(side="right", padx=(10,2))
        self._filter_var = tk.StringVar(value="الكل")
        self._filter_var.trace_add("write", lambda *a: self._filter())
        ttk.Combobox(fbar, textvariable=self._filter_var,
                     values=["الكل"] + self.STATUTS[:-1],
                     width=9, state="readonly").pack(side="right", padx=4)

        sbtn(fbar, "‏📤 Excel",        self._export_excel, SUCCESS,   pady=3, padx=6).pack(side="left", padx=4)
        sbtn(fbar, "‏🖨️ طباعة",       self._print_liste,  "#8e44ad", pady=3, padx=6).pack(side="left", padx=4)

        # ── جدول المترشحين ────────────────────────────────────────────────
        tbl = tk.Frame(self, bg=BG)
        tbl.pack(fill="both", expand=True, padx=16, pady=(0,2))

        cols = ["#", "الاسم واللقب", "التخصص", "نوع التكوين", "الحالة", "رقم الملف"]
        self.tree, sb = mk_tree(tbl, cols, h=16)
        for col, w in zip(cols, [40, 240, 120, 90, 80, 90]):
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col, width=w, minwidth=40,
                             anchor="center", stretch=(col=="الاسم واللقب"))
        for st, bg in self.TAG_BG.items():
            tag = st or "none"
            self.tree.tag_configure(tag, background=bg,
                                    foreground=self.COLORS.get(st, TEXT))
        self.tree.pack(side="right", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._show_context_menu)

        # ── قائمة السياق ──────────────────────────────────────────────────
        self._ctx_menu = tk.Menu(self, tearoff=0, bg=BG2, fg=TEXT,
                                 activebackground=ACCENT,
                                 activeforeground=WHITE, bd=0)
        for label, ico, clr, st in btn_defs:
            self._ctx_menu.add_command(
                label=f"{ico} {label}",
                command=lambda s=st: self._set_status(s))
        self._ctx_menu.add_separator()
        self._ctx_menu.add_command(label="📄 تفاصيل المترشح",
                                   command=self._show_details)

        # بيانات داخلية
        self._rows = []          # كل صفوف الدورة
        self._filtered = []      # الصفوف المعروضة بعد الفلترة

    # ─── تحميل بيانات الدورة ─────────────────────────────────────────────────
    def refresh(self):
        """تحديث قائمة الدورات في الـ combobox"""
        dawras = []
        for r in load_sheet("دورات"):
            if not any(r): continue
            raqm = fmt_raqm(str(r[2] or ""))
            if raqm: dawras.append(raqm)
        self.cb_dawra["values"] = dawras
        # إذا في دورة محددة نحافظ عليها أو نأخذ الأولى
        cur = self.raqm_var.get()
        if cur not in dawras and dawras:
            self.raqm_var.set(dawras[-1])
        self._load_dawra()

    def _load_dawra(self):
        """قراءة مترشحي الدورة المختارة وعرضهم"""
        raqm = self.raqm_var.get().strip()
        self._rows = []
        if not raqm:
            self._filter(); return

        for r in load_sheet("مترشحون"):
            if not any(r): continue
            n = len(r)
            r_raqm = fmt_raqm(str(r[17] if n > 17 else ""))
            if r_raqm != raqm: continue
            self._rows.append(r)

        self._filter()

    def _filter(self):
        """تطبيق الفلتر والبحث وتحديث الجدول"""
        filt  = self._filter_var.get()
        search = self._search_var.get().strip().lower()

        self._filtered = []
        for r in self._rows:
            n      = len(r)
            status = str(r[19] if n > 19 else "") or ""
            name   = str(r[1]  if n > 1  else "") or ""

            if filt != "الكل" and status != filt:
                continue
            if search and search not in name.lower():
                continue
            self._filtered.append(r)

        self._refresh_tree()
        self._refresh_stats()

    def _refresh_tree(self):
        """إعادة رسم الجدول"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, r in enumerate(self._filtered, 1):
            n      = len(r)
            name   = str(r[1]  if n > 1  else "") or "—"
            takh   = str(r[16] if n > 16 else "") or "—"
            naw    = str(r[18] if n > 18 else "") or "—"
            status = str(r[19] if n > 19 else "") or ""
            milaf  = str(r[26] if n > 26 else "") or "—"
            tag    = status or "none"
            self.tree.insert("", "end",
                             values=(i, name, takh, naw,
                                     status or "—", milaf),
                             tags=(tag,))

    def _refresh_stats(self):
        """تحديث بطاقات الإحصاء"""
        total = len(self._rows)
        counts = {s: 0 for s in self.STATUTS}
        for r in self._rows:
            n = len(r)
            st = str(r[19] if n > 19 else "") or ""
            counts[st] = counts.get(st, 0) + 1

        self._stat_labels["total"].config(text=str(total))
        for key, lbl in self._stat_labels.items():
            if key == "total": continue
            lbl.config(text=str(counts.get(key, 0)))

    # ─── تغيير الحالة ────────────────────────────────────────────────────────
    def _set_status(self, status):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("⚠️", "اختر مترشحاً أولاً"); return

        raqm = self.raqm_var.get().strip()
        if not raqm: return

        # تجميع الأسماء المحددة
        names = []
        for item in sel:
            vals = self.tree.item(item)["values"]
            if vals: names.append(str(vals[1]))

        if not names: return

        # تأكيد إذا كثر
        if len(names) > 3:
            if not messagebox.askyesno("تأكيد",
                f"تغيير حالة {len(names)} مترشح إلى «{status or 'بلا حالة'}»؟"):
                return

        # كتابة في Excel
        try:
            wb = openpyxl.load_workbook(EXCEL_FILE)
            ws = wb[SHEETS["مترشحون"]]
            today = datetime.now().strftime("%Y-%m-%d")
            changed = 0
            for row_i in range(2, ws.max_row + 1):
                cell_name = str(ws.cell(row_i, 2).value or "").strip()
                if cell_name not in names: continue
                ws.cell(row_i, 20).value = status  # حالة النجاح (عمود 20)
                # إذا ناجح وليس عنده رقم شهادة → نولّد
                if status == "ناجح":
                    if not ws.cell(row_i, 21).value:
                        ws.cell(row_i, 21).value = next_num(12)
                        ws.cell(row_i, 22).value = today
                    # تخصص مزدوج → شهادة ثانية
                    takh = str(ws.cell(row_i, 17).value or "")
                    if "البضائع" in takh and "الأشخاص" in takh:
                        if not ws.cell(row_i, 23).value:
                            ws.cell(row_i, 23).value = next_num(14)
                            ws.cell(row_i, 24).value = today
                changed += 1
            wb.save(EXCEL_FILE)
        except PermissionError:
            messagebox.showerror("❌ الملف مفتوح",
                                 "أغلق ملف Excel أولاً ثم حاول مجدداً"); return
        except Exception as e:
            messagebox.showerror("❌ خطأ", str(e)); return

        # تحديث العدد في ورقة الدورات
        try: _update_dawra_count_excel(raqm)
        except Exception: pass

        self._load_dawra()
        st_label = status or "بلا حالة"
        messagebox.showinfo("✅", f"تم تحديث {changed} مترشح → {st_label}")

    # ─── نافذة التفاصيل ──────────────────────────────────────────────────────
    def _show_details(self):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0])["values"]
        if not vals: return
        name = str(vals[1])
        for r in self._rows:
            if str(r[1] or "").strip() == name.strip():
                self._open_detail_window(r); return

    def _open_detail_window(self, r):
        n = len(r)
        win = tk.Toplevel(self)
        win.title("تفاصيل المترشح")
        win.configure(bg=BG); win.grab_set()
        win.geometry("480x520")
        tk.Label(win, text="📋  بيانات المترشح", font=_efont(13,bold=True),
                 bg=BG, fg=ACCENT2).pack(pady=(12,6))
        frm = tk.Frame(win, bg=BG2, padx=20, pady=14)
        frm.pack(fill="both", expand=True, padx=16, pady=4)

        fields = [
            ("الاسم واللقب",              1),
            ("اللقب بالفرنسية",           2),
            ("الاسم بالفرنسية",           3),
            ("تاريخ ومكان الازدياد",      4),
            ("العنوان",                   5),
            ("رقم التعريف الوطني",        6),
            ("رقم رخصة السياقة",          7),
            ("تاريخ صدور رخصة السياقة",   8),
            ("تخصص التكوين",              16),
            ("رقم الدورة",                17),
            ("نوع التكوين",               18),
            ("الحالة",                    19),
            ("رقم الشهادة",               20),
            ("تاريخ الشهادة",             21),
            ("رقم الملف",                 26),
        ]
        for i, (label, idx) in enumerate(fields):
            val = str(r[idx] if idx < n else "") or "—"
            color = self.COLORS.get(val, TEXT) if label == "الحالة" else TEXT
            tk.Label(frm, text=label + ":", bg=BG2, fg=TEXT2,
                     font=("Arial",9), anchor="e",
                     width=22).grid(row=i, column=1, sticky="e", pady=3, padx=5)
            tk.Label(frm, text=val, bg=BG2, fg=color,
                     font=("Arial",9,"bold"), anchor="w",
                     width=28).grid(row=i, column=0, sticky="w", pady=3)

        sbtn(win, "إغلاق", win.destroy, BG3, pady=6,
             padx=20).pack(pady=10)

    # ─── قائمة السياق (كليك يمين) ────────────────────────────────────────────
    def _show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _on_double_click(self, event):
        self._show_details()

    # ─── تصدير Excel ─────────────────────────────────────────────────────────
    def _export_excel(self):
        raqm = self.raqm_var.get().strip()
        if not self._rows:
            messagebox.showwarning("⚠️", "لا توجد بيانات للتصدير"); return
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile=f"تسيير_دورة_{raqm.replace('/','_')}.xlsx")
        if not path: return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "تسيير الدورة"
            ws.sheet_view.rightToLeft = True
            headers = ["#","الاسم واللقب","اللقب بالفرنسية","الاسم بالفرنسية",
                       "تاريخ ومكان الازدياد","رقم التعريف الوطني",
                       "التخصص","نوع التكوين","الحالة","رقم الملف"]
            ws.append(headers)
            for i, r in enumerate(self._filtered, 1):
                n = len(r)
                ws.append([
                    i,
                    str(r[1]  if n>1  else ""),
                    str(r[2]  if n>2  else ""),
                    str(r[3]  if n>3  else ""),
                    str(r[4]  if n>4  else ""),
                    str(r[6]  if n>6  else ""),
                    str(r[16] if n>16 else ""),
                    str(r[18] if n>18 else ""),
                    str(r[19] if n>19 else ""),
                    str(r[26] if n>26 else ""),
                ])
            wb.save(path)
            messagebox.showinfo("✅", f"تم تصدير {len(self._filtered)} مترشح")
        except Exception as e:
            messagebox.showerror("❌", str(e))

    # ─── طباعة القائمة ───────────────────────────────────────────────────────
    def _print_liste(self):
        raqm = self.raqm_var.get().strip()
        if not raqm:
            messagebox.showwarning("⚠️", "اختر دورة أولاً"); return
        try:
            import محاضر_الدورة as md
            path = md.export_liste_mtrsheen(raqm)
            if path:
                md.open_file(path)
        except Exception as e:
            messagebox.showerror("❌", str(e))

    def _mahdar_mudalawa(self):
        raqm = self.raqm_var.get().strip()
        if not raqm:
            messagebox.showwarning("⚠️", "اختر دورة أولاً"); return
        try:
            import محاضر_الدورة as md
            mahdar = ""
            for r in load_sheet("دورات"):
                if not any(r): continue
                if fmt_raqm(str(r[2] or "")) == raqm:
                    mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                    break
            path = md.fill_mahdar_mudalawa(raqm, raqm_mahdar=mahdar if mahdar else None)
            if path:
                _set_dawra_hala(raqm, "منتهية")
                md.open_file(path)
        except Exception as e:
            messagebox.showerror("❌", str(e))

    def _mahdar_fath(self):
        raqm = self.raqm_var.get().strip()
        if not raqm:
            messagebox.showwarning("⚠️", "اختر دورة أولاً"); return
        try:
            import محاضر_الدورة as md
            mahdar = ""
            for r in load_sheet("دورات"):
                if not any(r): continue
                if fmt_raqm(str(r[2] or "")) == raqm:
                    mahdar = str(r[9] or "").strip() if len(r) > 9 else ""
                    break
            path = md.fill_mahdar_fath(raqm, raqm_mahdar=mahdar if mahdar else None)
            if path:
                _set_dawra_hala(raqm, "جارية")
                md.open_file(path)
        except Exception as e:
            messagebox.showerror("❌", str(e))

    def _liste_najihin(self):
        raqm = self.raqm_var.get().strip()
        if not raqm:
            messagebox.showwarning("⚠️", "اختر دورة أولاً"); return
        try:
            import محاضر_الدورة as md
            path = md.export_liste_najihin(raqm)
            if path:
                md.open_file(path)
        except Exception as e:
            messagebox.showerror("❌", str(e))


# ══════════════════════════════════════════════════════════════════════════════
# صفحة تسيير الدورة — DawraGestionPage
# ══════════════════════════════════════════════════════════════════════════════


if __name__=="__main__":
    try:
        def _install(name):
            for args in [
                [sys.executable,"-m","pip","install",name,"-q"],
                [sys.executable,"-m","pip","install",name,"-q","--user"],
            ]:
                try:
                    subprocess.check_call(args, stdout=subprocess.DEVNULL,
                                          stderr=subprocess.DEVNULL)
                    return True
                except Exception:
                    continue
            return False

        for pkg in ["dateutil","openpyxl","reportlab","pymupdf","arabic_reshaper","bidi"]:
            try: __import__(pkg if pkg!="dateutil" else "dateutil.relativedelta")
            except ImportError:
                name="python-dateutil" if pkg=="dateutil" else pkg
                _install(name)

        # ── التحقق من التحديثات ────────────────────────────────────────
        try:
            import updater
            updated = updater.check_and_update(silent=True)
            if updated:
                updater.restart()
        except Exception:
            pass  # لا نوقف البرنامج إذا فشل التحديث

        app=App(); app.mainloop()
    except Exception as e:
        log_error(e)
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("خطأ في التشغيل", traceback.format_exc())
        root.destroy()