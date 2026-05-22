#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
وحدة طباعة شهادات الكفاءة المهنية
"""
import os, sys, platform, subprocess
from datetime import datetime
from dateutil.relativedelta import relativedelta
from utils import fmt_date as _fmt_date_utils, fmt_raqm, clean_wilaya, clean_takh, EXCEL_FILE, SHEETS, COLS_MARKAZ

FONT_NAME      = "ArabicFont"
FONT_NAME_BOLD = "ArabicFontBold"
_OFFSET_Y = 0


def _get_font_path(bold=False):
    if platform.system() == "Windows":
        if bold:
            for f in [r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\tahomabd.ttf",
                      r"C:\Windows\Fonts\calibrib.ttf"]:
                if os.path.exists(f): return f
        for f in [r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\tahoma.ttf",
                  r"C:\Windows\Fonts\calibri.ttf", r"C:\Windows\Fonts\arialbd.ttf"]:
            if os.path.exists(f): return f
    else:
        for f in ["/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
                  "/usr/share/fonts/truetype/freefont/FreeSans.ttf"]:
            if os.path.exists(f): return f
    return None


def _register_font():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try: pdfmetrics.getFont(FONT_NAME); return
    except Exception: pass
    path = _get_font_path(bold=False)
    if path: pdfmetrics.registerFont(TTFont(FONT_NAME, path))
    else: raise RuntimeError("لم يتم العثور على خط مناسب")
    # خط سميك
    try: pdfmetrics.getFont(FONT_NAME_BOLD)
    except Exception:
        path_bold = _get_font_path(bold=True)
        if path_bold:
            pdfmetrics.registerFont(TTFont(FONT_NAME_BOLD, path_bold))


EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "بيانات_مركز_التكوين.xlsx")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "شهادات_PDF")
SHEETS     = {"مركز": "مركز_التكوين", "مترشحون": "المترشحون"}
COLS_MARKAZ = ["اسم مركز التكوين","رقم الاعتماد","تاريخ الاعتماد","عدد الأقسام",
               "مديرية النقل لولاية","عنوان مركز التكوين","بلدية المركز","ممثل مركز التكوين",
               "المسؤول البيداغوجي","الأساتذة","رقم الهاتف","البريد الإلكتروني"]

# الإحداثيات الأصلية
FIELDS = {
    "wilaya":    (11.8, 27.0, 10),
    "docnum":    (22.2, 17.0, 10),
    "director":  (56.1, 33.4, 10),
    "name":      (59.4, 22.8, 11),
    "dob":       (63.2, 24.4, 10),
    "address":   (66.5, 23.9, 10),
    "vehicles":  (69.9, 38.0,  9),
    "date_from": (69.9, 62.0,  9),
    "date_to":   (69.9, 78.0,  9),
    "city":      (76.7, 44.0, 10),
    "sign_date": (76.7, 70.0, 10),
}


def ar(text):
    """تحويل النص للعرض الصحيح — مع الحفاظ على اتجاه التواريخ والأرقام"""
    import arabic_reshaper
    from bidi.algorithm import get_display
    import re
    s = str(text).strip()
    if not s: return s

    # تاريخ فقط YYYY/MM/DD
    if re.match(r'^\d{4}/\d{2}/\d{2}$', s):
        return s

    # تاريخ مع مكان — بأي صيغة وبأي فاصل
    date_match = re.match(r'^(\d{4}/\d{2}/\d{2})\s*[-–]\s*(.+)$', s)
    if not date_match:
        # محاولة الصيغة القديمة DD/MM/YYYY-مكان
        date_match2 = re.match(r'^(\d{2}/\d{2}/\d{4})\s*[-–]?\s*(.+)$', s)
        if date_match2:
            from datetime import datetime as _dt
            try:
                d = _dt.strptime(date_match2.group(1), "%d/%m/%Y").strftime("%Y/%m/%d")
                place_part = get_display(arabic_reshaper.reshape(date_match2.group(2).strip()))
                return f"{place_part} - {d}"
            except Exception: pass
    if date_match:
        date_part  = date_match.group(1)
        place_part = get_display(arabic_reshaper.reshape(date_match.group(2).strip()))
        return f"{place_part} - {date_part}"

    # نص عربي عادي
    return get_display(arabic_reshaper.reshape(s))


def _parse_date(ds):
    """تحويل أي صيغة تاريخ إلى كائن datetime — يدعم YYYY/MM/DD و YYYY-MM-DD و DD/MM/YYYY"""
    import re as _re
    s = str(ds).strip()
    if not s or s.lower() == "none": return None
    # كائن datetime مباشرة
    if hasattr(ds, 'strftime'): return ds
    # YYYY/MM/DD
    if _re.match(r'^\d{4}/\d{2}/\d{2}$', s):
        try: return datetime.strptime(s, "%Y/%m/%d")
        except Exception: pass
    # YYYY-MM-DD
    try: return datetime.strptime(s.split(" ")[0], "%Y-%m-%d")
    except Exception: pass
    # DD/MM/YYYY
    m = _re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if m:
        try: return datetime.strptime(s, "%d/%m/%Y")
        except Exception: pass
    return None


def calc_expiry(ds):
    """حساب تاريخ انتهاء الشهادة (5 سنوات) — يقبل YYYY/MM/DD و YYYY-MM-DD و DD/MM/YYYY"""
    if not ds: return ""
    d = _parse_date(ds)
    if d: return (d + relativedelta(years=5)).strftime("%Y/%m/%d")
    return str(ds)


def format_date(ds):
    """تحويل أي صيغة تاريخ → YYYY/MM/DD"""
    if not ds: return ""
    d = _parse_date(ds)
    if d: return d.strftime("%Y/%m/%d")
    return str(ds).strip()


def draw_field(cv, text, top_pct, right_pct, font_size=10,
               color=(0, 0, 0.45), center=False, bold=False):
    if not text: return
    from reportlab.lib.pagesizes import A5
    W, H    = A5
    x_right = W * (1 - right_pct / 100)
    y       = H * (1 - top_pct  / 100) - 6 + _OFFSET_Y
    font    = FONT_NAME_BOLD if bold else FONT_NAME
    try: cv.setFont(font, font_size)
    except Exception: cv.setFont(FONT_NAME, font_size)
    cv.setFillColorRGB(*color)
    if center:
        width_pt = W * 0.14
        x_center = x_right - width_pt / 2
        cv.drawCentredString(x_center, y, ar(str(text)))
    else:
        cv.drawRightString(x_right, y, ar(str(text)))


def load_markaz():
    """قراءة بيانات المركز — يقرأ رؤوس الأعمدة من Excel مباشرة لتجنب أي إزاحة"""
    if not os.path.exists(EXCEL_FILE): return {}
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEETS["مركز"]]
    cols = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
    for row in ws.iter_rows(min_row=2, max_row=2, values_only=True):
        return {k: (str(v) if v is not None else "") for k, v in zip(cols, row) if k}
    return {}


def load_dawra_info(raqm_dawra):
    if not os.path.exists(EXCEL_FILE) or not raqm_dawra: return {}
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE)
    try:
        ws = wb["الدورات"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[2] and str(row[2]) == str(raqm_dawra):
                return {"date_ouv": format_date(str(row[0] or "")),
                        "date_clo": format_date(str(row[1] or "")),
                        "raqm": str(row[2] or "")}
    except Exception: pass
    return {}


def load_najihin(raqm_dawra=None):
    """قراءة الناجحين من Excel — يقرأ رؤوس الأعمدة تلقائياً لدعم أي هيكل"""
    if not os.path.exists(EXCEL_FILE): return []
    import openpyxl
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEETS["مترشحون"]]
    headers = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
    results = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        d = {k: (str(v) if v is not None else "") for k, v in zip(headers, row) if k}
        if d.get("حالة النجاح") != "ناجح": continue
        if raqm_dawra and d.get("رقم الدورة","").strip() != str(raqm_dawra).strip(): continue
        for key in ["تاريخ الشهادة", "تاريخ الشهادة 2", "تاريخ التسجيل", "تاريخ صدور رخصة السياقة"]:
            if key in d and d[key]:
                d[key] = format_date(d[key])
        results.append(d)
    return results


def generate_qr(data_dict):
    import qrcode
    text = (
        "N: "          + str(data_dict.get("docnum",     "")) + "\n"
        "Nom: "        + str(data_dict.get("name",       "")) + "\n"
        "Specialite: " + str(data_dict.get("vehicles",   "")) + "\n"
        "Wilaya: "     + str(data_dict.get("wilaya",     "")) + "\n"
        "Centre: "     + str(data_dict.get("centre",     "")) + "\n"
        "N.Session: "  + str(data_dict.get("raqm_dawra", "")) + "\n"
        "Du: "         + str(data_dict.get("date_ouv",   "")) + "\n"
        "Au: "         + str(data_dict.get("date_clo",   "")) + "\n"
        "Valide: "     + str(data_dict.get("date_from",  ""))
        + " - "        + str(data_dict.get("date_to",    ""))
    )
    qr = qrcode.QRCode(version=3,
                       error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=4, border=1)
    qr.add_data(text); qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")


def draw_qr(cv, data_dict):
    """QR Code — أسفل يمين الشهادة، مرفوع 1.5سم عن الهامش السفلي"""
    try:
        import io
        from reportlab.lib.pagesizes import A5
        from reportlab.lib.utils import ImageReader
        W, H    = A5
        qr_size = 2.2 / 2.54 * 72          # 2.2سم بالنقاط
        margin  = 1.5 / 2.54 * 72          # 1.5سم من الحافة اليمنى والسفلى
        x = W - qr_size - margin
        y = margin + _OFFSET_Y
        buf = io.BytesIO()
        generate_qr(data_dict).save(buf, format="PNG")
        buf.seek(0)
        from reportlab.lib.utils import ImageReader
        cv.drawImage(ImageReader(buf), x, y, qr_size, qr_size)
    except Exception: pass


def _clean_wilaya(w):
    """حذف الرقم والشرطة من اسم الولاية — '30-ورقلة' → 'ورقلة'"""
    import re
    return re.sub(r'^\d+-', '', str(w or '')).strip()


def _clean_takh(takh):
    return (str(takh)
            .replace("نقل الأشخاص و البضائع", "أشخاص/بضائع")
            .replace("نقل الأشخاص",            "أشخاص")
            .replace("نقل البضائع",             "بضائع")
            .replace("نقل المواد الخطرة",       "مواد خطرة")
            .replace("الأشخاص و البضائع",       "أشخاص/بضائع")
            .replace("الأشخاص", "أشخاص")
            .replace("البضائع", "بضائع")
            .replace("المواد الخطرة", "مواد خطرة"))


def print_dawra(raqm_dawra=None, output_dir=None):
    global _OFFSET_Y
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A5

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    _register_font()

    markaz     = load_markaz()
    najihin    = load_najihin(raqm_dawra)
    dawra_info = load_dawra_info(raqm_dawra) if raqm_dawra else {}

    if not najihin:
        print("لا يوجد ناجحون"); return None

    W, H  = A5
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    raqm  = str(raqm_dawra) if raqm_dawra else "الكل"
    out   = os.path.join(out_dir, f"شهادات_دورة_{raqm}_{ts}.pdf")

    cv = rl_canvas.Canvas(out, pagesize=A5, invariant=1)
    cv.setPageSize(A5)
    # تأكيد حجم الورق A5 في metadata الـ PDF لمنع الطابعة من التكبير التلقائي
    try:
        cv._doc.pagesize = A5
    except Exception:
        pass
    _OFFSET_Y = 0
    page = 0

    for mtrsh in najihin:
        takh      = str(mtrsh.get("تخصص التكوين","") or "")
        takh_norm = takh.replace("نقل الأشخاص و البضائع","الأشخاص و البضائع").replace("نقل ","")
        sign_raw  = str(mtrsh.get("تاريخ الشهادة","") or datetime.now().strftime("%Y/%m/%d"))
        num_sh    = str(mtrsh.get("رقم الشهادة","") or "")

        shahadas = [(takh_norm if takh_norm != "الأشخاص و البضائع" else "الأشخاص", num_sh, sign_raw)]
        if takh_norm == "الأشخاص و البضائع":
            shahadas.append(("البضائع",
                             str(mtrsh.get("رقم الشهادة 2","") or ""),
                             str(mtrsh.get("تاريخ الشهادة 2","") or sign_raw)))

        for takh_sh, ns, sd in shahadas:
            if page > 0:
                cv.showPage()
                cv.setPageSize(A5)
            page += 1

            data = {
                "wilaya":    _clean_wilaya(markaz.get("مديرية النقل لولاية","")),
                "docnum":    ns,
                "director":  _clean_wilaya(markaz.get("مديرية النقل لولاية","")),
                "name":      str(mtrsh.get("الاسم واللقب","") or ""),
                "dob":       str(mtrsh.get("تاريخ ومكان الازدياد","") or ""),
                "address":   str(mtrsh.get("العنوان","") or ""),
                "vehicles":  _clean_takh(takh_sh),
                "date_from": format_date(sd),
                "date_to":   calc_expiry(sd),
                "city":      _clean_wilaya(markaz.get("مديرية النقل لولاية","")),
                "sign_date": format_date(sd),
            }

            for fid, (top, right, fsize) in FIELDS.items():
                is_bold = fid in ("name","dob","docnum","vehicles","date_from","date_to","sign_date","city")
                draw_field(cv, data.get(fid,""), top, right, fsize,
                           center=(fid == "vehicles"), bold=is_bold)

            qr_data = dict(data)
            qr_data["centre"]     = str(markaz.get("اسم مركز التكوين","") or "")
            qr_data["raqm_dawra"] = str(raqm_dawra or "")
            qr_data["date_ouv"]   = dawra_info.get("date_ouv","")
            qr_data["date_clo"]   = dawra_info.get("date_clo","")
            draw_qr(cv, qr_data)

    # منع الطابعة من تكبير الصفحة — يحدد A5 صراحةً في ملف PDF
    try:
        cv._doc.info.subject  = "A5"
        cv._doc.info.keywords = "PageSize:A5"
    except Exception:
        pass
    cv.save()
    print(f"{page} شهادة → {out}")
    return out


def _generate_qr_generic(text):
    """توليد QR Code نصي عام"""
    import qrcode
    qr = qrcode.QRCode(version=3,
                       error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=4, border=1)
    qr.add_data(text); qr.make(fit=True)
    return qr.make_image(fill_color="black", back_color="white")


def _draw_qr_generic(cv, text, x, y, size_cm=2.5):
    """رسم QR Code في موضع محدد"""
    try:
        import io
        from reportlab.lib.utils import ImageReader
        from reportlab.lib.units import cm
        size = size_cm * cm
        buf = io.BytesIO()
        _generate_qr_generic(text).save(buf, format="PNG")
        buf.seek(0)
        cv.drawImage(ImageReader(buf), x, y, size, size)
    except Exception:
        pass


def print_tasjil_wasl_combined(mtrsh, montant_paye, raqm_wasl,
                               montant_paye_precedent=0.0, output_dir=None):
    """
    طباعة شهادة التسجيل + وصل التسديد في ورقة A4 أفقية (landscape).
    يمين الورقة = شهادة التسجيل
    يسار الورقة = وصل التسديد
    خط منقط عمودي في المنتصف للقص.
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import cm

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    _register_font()

    markaz     = load_markaz()
    W, H       = landscape(A4)
    HALF       = W / 2
    ts         = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom        = str(mtrsh.get("الاسم واللقب","") or "متربص").replace(" ","_")
    raqm_safe  = str(raqm_wasl).replace("/","_")
    out        = os.path.join(out_dir, f"تسجيل_وصل_{raqm_safe}_{nom}_{ts}.pdf")

    takh       = _clean_takh(str(mtrsh.get("تخصص التكوين","") or ""))
    dawra      = str(mtrsh.get("رقم الدورة","") or "")
    tasjil_d   = format_date(str(mtrsh.get("تاريخ التسجيل","") or datetime.now().strftime("%Y-%m-%d")))
    total      = get_montant(takh)
    paye_avant = float(montant_paye_precedent or 0)
    paye_now   = float(montant_paye or 0)
    total_paye = paye_avant + paye_now
    reste      = total - total_paye
    ville      = _clean_wilaya(markaz.get("مديرية النقل لولاية",""))
    today      = datetime.now().strftime("%Y/%m/%d")
    m          = 0.7*cm   # هامش داخلي

    cv = rl_canvas.Canvas(out, pagesize=landscape(A4))
    cv.setPageSize(landscape(A4))

    # ── خط فاصل منقط عمودي ──────────────────────────────────────
    cv.setStrokeColorRGB(0.5,0.5,0.5); cv.setLineWidth(0.5); cv.setDash(4,4)
    cv.line(HALF, 0.5*cm, HALF, H-0.5*cm)
    cv.setDash()
    cv.setFont(FONT_NAME, 7); cv.setFillColorRGB(0.6,0.6,0.6)
    cv.drawCentredString(HALF, H/2, ar('✂'))

    # ── دوال مساعدة ─────────────────────────────────────────────
    def wr(text, x0, x1, y, size=8, center=False, gap=3, color=(0,0,0)):
        cv.setFont(FONT_NAME, size); cv.setFillColorRGB(*color)
        if center: cv.drawCentredString((x0+x1)/2, y, ar(str(text)))
        else: cv.drawRightString(x1-m, y, ar(str(text)))
        return y - size*0.4 - gap

    def hline(x0, x1, y, w=0.5):
        cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(w)
        cv.line(x0+m, y, x1-m, y)

    def section_frame(x0, x1):
        cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(1.2)
        cv.rect(x0+0.5*cm, 0.5*cm, x1-x0-1*cm, H-1*cm)
        cv.setLineWidth(0.3)
        cv.rect(x0+0.7*cm, 0.7*cm, x1-x0-1.4*cm, H-1.4*cm)

    def header(x0, x1, y):
        y = wr('الجمهورية الجزائرية الديمقراطية الشعبية', x0,x1,y, 9, center=True, gap=3)
        y = wr('وزارة الداخلية والجماعات المحلية والنقل', x0,x1,y, 8, center=True, gap=3)
        hline(x0,x1,y); y -= 5
        y = wr(f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}', x0,x1,y, 8, center=True, gap=2)
        return y

    # ════════════════════════════════════════════════════════════
    # يمين: شهادة التسجيل — مطابقة للنموذج الرسمي
    # ════════════════════════════════════════════════════════════
    x0, x1 = HALF, W
    section_frame(x0, x1)
    y = H - 1.3*cm

    # رأس
    y = wr('الجمهورية الجزائرية الديمقراطية الشعبية', x0,x1,y, 9, center=True, gap=3)
    y = wr('وزارة الداخلية والجماعات المحلية والنقل', x0,x1,y, 8, center=True, gap=4)
    hline(x0,x1,y); y -= 5

    # اسم المركز + التخصصات
    y = wr(f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}    للحصول على شهادة الكفاءة المهنية', x0,x1,y, 8, center=True, gap=3)
    y = wr('لنقل الأشخاص و البضائع و/أو نقل المواد الخطرة', x0,x1,y, 8, center=True, gap=3)
    y = wr(f'رقم الهاتف : {markaz.get("رقم الهاتف","")}', x0,x1,y, 8, gap=3)
    y = wr(f'البريد الإلكتروني : {markaz.get("البريد الإلكتروني","")}', x0,x1,y, 8, gap=4)
    hline(x0,x1,y); y -= 10

    # عنوان الوثيقة
    cv.setFont(FONT_NAME,13); cv.setFillColorRGB(0,0,0)
    cv.drawCentredString((x0+x1)/2, y, ar('شهادة التسجيل')); y -= 14

    # النص التمهيدي
    y = wr(f'يشهد ممثل مركز التكوين : {markaz.get("ممثل مركز التكوين","")} ، لنيل شهادة الكفاءة المهنية لنقل الأشخاص', x0,x1,y, 8, gap=3)
    y = wr('و البضائع و/أو نقل المواد الخطرة قبول ملف المتربص أدناه :', x0,x1,y, 8, gap=6)

    # بيانات المتربص
    for label, val in [
        ('الاسم واللقب',              mtrsh.get("الاسم واللقب","")),
        ('تاريخ ومكان الازدياد',      mtrsh.get("تاريخ ومكان الازدياد","")),
        ('رقم التعريف الوطني',        mtrsh.get("رقم التعريف الوطني","")),
        ('رقم رخصة السياقة',          mtrsh.get("رقم رخصة السياقة","")),
        ('تاريخ صدور رخصة السياقة',   format_date(str(mtrsh.get("تاريخ صدور رخصة السياقة","") or ""))),
        ('العنوان',                   mtrsh.get("العنوان","")),
        ('رقم الهاتف',                mtrsh.get("رقم هاتف المتربص","")),
        ('البريد الإلكتروني',          mtrsh.get("البريد الإلكتروني","")),
    ]:
        y = wr(f'- {label} : {val}', x0,x1,y, 8, gap=4)

    y -= 4
    # سطر التخصص والدورة والتاريخ في سطر واحد
    y = wr(f'مسجل لدينا في تخصص :  {takh}     ضمن الدورة رقم :  {dawra}     تاريخ التسجيل :  {tasjil_d}', x0,x1,y, 8, gap=6)

    # النص الختامي
    y = wr('سُلِّمت هذه الشهادة للإدلاء بها في حدود ما يسمح به القانون.', x0,x1,y, 8, gap=3)
    y = wr('ملاحظة : هذه الشهادة صالحة لمدة ثلاثة (03) أشهر.', x0,x1,y, 8, gap=12)

    # التوقيع
    cv.setFont(FONT_NAME,8); cv.setFillColorRGB(0,0,0)
    cv.drawRightString(x1-m, y, ar(f'حُررت بـ {ville} في : {today}')); y -= 20
    cv.drawRightString(x1-m, y, ar('ختم وتوقيع ممثل مركز التكوين'))

    # QR
    qr1 = (f"نوع الوثيقة: شهادة تسجيل\nالاسم: {mtrsh.get('الاسم واللقب','')}\n"
           f"رقم التعريف: {mtrsh.get('رقم التعريف الوطني','')}\nالتخصص: {takh}\n"
           f"الدورة: {dawra}\nالتسجيل: {tasjil_d}\nالمركز: {markaz.get('اسم مركز التكوين','')}")
    _draw_qr_generic(cv, qr1, x0+m, 0.7*cm, size_cm=2.0)

    # ════════════════════════════════════════════════════════════
    # يسار: وصل التسديد
    # ════════════════════════════════════════════════════════════
    x0, x1 = 0, HALF
    section_frame(x0, x1)
    y = H - 1.3*cm
    y = header(x0, x1, y)
    y = wr(f'رقم الهاتف : {markaz.get("رقم الهاتف","")}', x0,x1,y, 7, center=True, gap=2)
    hline(x0,x1,y); y -= 8
    cv.setFont(FONT_NAME,13); cv.setFillColorRGB(0,0,0)
    cv.drawCentredString((x0+x1)/2, y, ar('وصل تسديد حقوق التكوين')); y -= 6
    cv.setFont(FONT_NAME,9)
    cv.drawCentredString((x0+x1)/2, y, ar(f'رقم الوصل : {raqm_wasl}')); y -= 12
    y = wr(f'تسلّمنا من السيد/السيدة : {mtrsh.get("الاسم واللقب","")}', x0,x1,y, 9, gap=4)
    y = wr(f'المسجل في تخصص : {takh}     —     الدورة رقم : {dawra}', x0,x1,y, 8, gap=8)
    # جدول المبالغ
    tbl_rows = [
        ('المبلغ الإجمالي لحقوق التكوين', f'{total:,.2f} دج'),
        ('المسدَّد سابقاً',               f'{paye_avant:,.2f} دج'),
        ('المسدَّد حالياً',               f'{paye_now:,.2f} دج'),
        ('مجموع ما تم تسديده',            f'{total_paye:,.2f} دج'),
        ('باقي الدين',                    f'{reste:,.2f} دج'),
    ]
    rh = 0.55*cm; ty = y
    for i, (l, v) in enumerate(tbl_rows):
        if i%2==0:
            cv.setFillColorRGB(0.94,0.94,0.94)
            cv.rect(x0+m, ty-rh+2, x1-x0-2*m, rh, fill=1, stroke=0)
        c = (0.7,0,0) if (i==4 and reste>0) else (0,0,0)
        cv.setFont(FONT_NAME,8); cv.setFillColorRGB(*c)
        cv.drawRightString(x1-m, ty-4, ar(l))
        cv.drawString(x0+m+4, ty-4, ar(v))
        cv.setStrokeColorRGB(0.8,0.8,0.8); cv.setLineWidth(0.3)
        cv.line(x0+m, ty-rh+2, x1-m, ty-rh+2)
        ty -= rh
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(0.5)
    cv.rect(x0+m, ty, x1-x0-2*m, y-ty)
    y = ty - 8
    y = wr('طريقة الدفع :  ☐ نقداً    ☐ حوالة بريدية    ☐ تحويل بنكي', x0,x1,y, 8, gap=10)
    cv.setFont(FONT_NAME,8); cv.setFillColorRGB(0,0,0)
    cv.drawString(x0+m, y, ar('توقيع المتربص :'))
    cv.drawRightString(x1-m, y, ar(f'حُررت بـ {ville} في : {today}')); y -= 20
    cv.drawRightString(x1-m, y, ar('ختم وتوقيع ممثل مركز التكوين'))
    qr2 = (f"نوع الوثيقة: وصل تسديد\nرقم الوصل: {raqm_wasl}\n"
           f"الاسم: {mtrsh.get('الاسم واللقب','')}\nالتخصص: {takh}\n"
           f"الإجمالي: {total:,.2f} دج\nالمسدَّد: {total_paye:,.2f} دج\n"
           f"الباقي: {reste:,.2f} دج\nالتاريخ: {today}")
    _draw_qr_generic(cv, qr2, x0+m, 0.7*cm, size_cm=2.0)

    cv.save()
    print(f"تسجيل + وصل (landscape) → {out}")
    return out



def print_shahada_tasjil(mtrsh, output_dir=None):
    """
    طباعة شهادة التسجيل لمتربص واحد — صيغة A4.
    mtrsh : قاموس بيانات المتربص من Excel
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    _register_font()

    markaz = load_markaz()
    W, H   = A4
    ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom    = str(mtrsh.get("الاسم واللقب","") or "متربص").replace(" ","_")
    out    = os.path.join(out_dir, f"شهادة_تسجيل_{nom}_{ts}.pdf")

    cv = rl_canvas.Canvas(out, pagesize=A4)
    cv.setPageSize(A4)

    # ── إطار مزدوج ───────────────────────────────────────────────
    cv.setStrokeColorRGB(0, 0, 0)
    cv.setLineWidth(1.5)
    cv.rect(1.5*cm, 1.5*cm, W-3*cm, H-3*cm)
    cv.setLineWidth(0.5)
    cv.rect(1.9*cm, 1.9*cm, W-3.8*cm, H-3.8*cm)

    y = H - 2.6*cm

    def wr(text, size=10, center=False, gap=5, color=(0,0,0)):
        nonlocal y
        cv.setFont(FONT_NAME, size)
        cv.setFillColorRGB(*color)
        if center:
            cv.drawCentredString(W/2, y, ar(str(text)))
        else:
            cv.drawRightString(W - 2.3*cm, y, ar(str(text)))
        y -= size * 0.45 + gap

    # ── رأس الوثيقة ──────────────────────────────────────────────
    wr('الجمهورية الجزائرية الديمقراطية الشعبية', 12, center=True, gap=5)
    wr('وزارة الداخلية والجماعات المحلية والنقل', 11, center=True, gap=6)
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(0.8)
    cv.line(2.2*cm, y, W-2.2*cm, y); y -= 8
    wr(f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}    للحصول على شهادة الكفاءة المهنية', 10, center=True, gap=4)
    wr('لنقل الأشخاص و البضائع و/أو نقل المواد الخطرة', 9, center=True, gap=4)
    wr(f'رقم الهاتف : {markaz.get("رقم الهاتف","")}', 9, gap=3)
    wr(f'البريد الإلكتروني : {markaz.get("البريد الإلكتروني","")}', 9, gap=4)
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(0.8)
    cv.line(2.2*cm, y, W-2.2*cm, y); y -= 14

    # ── العنوان ──────────────────────────────────────────────────
    cv.setFont(FONT_NAME, 15)
    cv.setFillColorRGB(0,0,0)
    cv.drawCentredString(W/2, y, ar('شهادة التسجيل'))
    y -= 20

    # ── النص التمهيدي ────────────────────────────────────────────
    wr(f'يشهد ممثل مركز التكوين : {markaz.get("ممثل مركز التكوين","")} ، لنيل شهادة الكفاءة المهنية لنقل الأشخاص', 10, gap=5)
    wr('و البضائع و/أو نقل المواد الخطرة قبول ملف المتربص أدناه :', 10, gap=10)

    # ── بيانات المتربص ───────────────────────────────────────────
    rows = [
        ('الاسم واللقب',               mtrsh.get("الاسم واللقب","")),
        ('تاريخ ومكان الازدياد',        mtrsh.get("تاريخ ومكان الازدياد","")),
        ('رقم التعريف الوطني',          mtrsh.get("رقم التعريف الوطني","")),
        ('رقم رخصة السياقة',            mtrsh.get("رقم رخصة السياقة","")),
        ('تاريخ صدور رخصة السياقة',     format_date(str(mtrsh.get("تاريخ صدور رخصة السياقة","") or ""))),
        ('العنوان',                     mtrsh.get("العنوان","")),
        ('رقم الهاتف',                  mtrsh.get("رقم هاتف المتربص","")),
        ('البريد الإلكتروني',            mtrsh.get("البريد الإلكتروني","")),
    ]
    for label, val in rows:
        wr(f'- {label} : {val}', 10, gap=6)

    y -= 6
    takh   = _clean_takh(str(mtrsh.get("تخصص التكوين","") or ""))
    dawra  = str(mtrsh.get("رقم الدورة","") or "")
    tasjil = format_date(str(mtrsh.get("تاريخ التسجيل","") or datetime.now().strftime("%Y-%m-%d")))
    wr(f'مسجل لدينا في تخصص :  {takh}     ضمن الدورة رقم :  {dawra}     تاريخ التسجيل :  {tasjil}', 10, gap=14)
    wr('سُلِّمت هذه الشهادة للإدلاء بها في حدود ما يسمح به القانون.', 10, gap=4)
    wr('ملاحظة : هذه الشهادة صالحة لمدة ثلاثة (03) أشهر.', 10, gap=20)

    # ── التوقيع ──────────────────────────────────────────────────
    ville = _clean_wilaya(markaz.get("مديرية النقل لولاية",""))
    today = datetime.now().strftime("%Y/%m/%d")
    wr(f'حُررت بـ {ville} في : {today}', 10, gap=25)
    wr('ختم وتوقيع ممثل مركز التكوين', 10, gap=0)

    # ── QR Code ──────────────────────────────────────────────────
    qr_text = (
        "نوع الوثيقة: شهادة تسجيل\n"
        f"الاسم واللقب: {mtrsh.get('الاسم واللقب','')}\n"
        f"رقم التعريف الوطني: {mtrsh.get('رقم التعريف الوطني','')}\n"
        f"التخصص: {takh}\n"
        f"رقم الدورة: {dawra}\n"
        f"تاريخ التسجيل: {tasjil}\n"
        f"المركز: {markaz.get('اسم مركز التكوين','')}"
    )
    _draw_qr_generic(cv, qr_text, 2*cm, 2*cm, size_cm=2.5)

    cv.save()
    print(f"شهادة تسجيل → {out}")
    return out


# ── مبالغ التكوين حسب التخصص ─────────────────────────────────────
MONTANTS_FORMATION = {
    "أشخاص/بضائع":      32300.00,   # يجب أن يأتي أولاً لتجنب التداخل مع "أشخاص"
    "أشخاص":            17000.00,
    "بضائع":            22500.00,
    "مواد خطرة":        16000.00,
}

def get_montant(takh):
    """إرجاع المبلغ الإجمالي حسب التخصص"""
    t = _clean_takh(str(takh or ""))
    for key, val in MONTANTS_FORMATION.items():
        if key in t:
            return val
    return 0.0


def print_wasl_tasdid(mtrsh, montant_paye, raqm_wasl,
                      montant_paye_precedent=0.0, output_dir=None):
    """
    طباعة وصل تسديد حقوق التكوين — صيغة A4.

    المعاملات:
        mtrsh                : قاموس بيانات المتربص
        montant_paye         : المبلغ المسدَّد حالياً
        raqm_wasl            : رقم الوصل التسلسلي
        montant_paye_precedent: مجموع ما سُدِّد سابقاً
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    _register_font()

    markaz = load_markaz()
    W, H   = A4
    ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom    = str(mtrsh.get("الاسم واللقب","") or "متربص").replace(" ","_")
    raqm_safe = str(raqm_wasl).replace("/","_")
    out    = os.path.join(out_dir, f"وصل_تسديد_{raqm_safe}_{nom}_{ts}.pdf")

    takh        = _clean_takh(str(mtrsh.get("تخصص التكوين","") or ""))
    total       = get_montant(takh)
    paye_avant  = float(montant_paye_precedent or 0)
    paye_now    = float(montant_paye or 0)
    total_paye  = paye_avant + paye_now
    reste       = total - total_paye
    dawra       = str(mtrsh.get("رقم الدورة","") or "")
    today       = datetime.now().strftime("%Y/%m/%d")
    ville       = _clean_wilaya(markaz.get("مديرية النقل لولاية",""))

    cv = rl_canvas.Canvas(out, pagesize=A4)
    cv.setPageSize(A4)

    # ── إطار مزدوج ───────────────────────────────────────────────
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(1.5)
    cv.rect(1.5*cm, 1.5*cm, W-3*cm, H-3*cm)
    cv.setLineWidth(0.5)
    cv.rect(1.9*cm, 1.9*cm, W-3.8*cm, H-3.8*cm)

    y = H - 2.6*cm

    def wr(text, size=10, center=False, gap=5, color=(0,0,0)):
        nonlocal y
        cv.setFont(FONT_NAME, size)
        cv.setFillColorRGB(*color)
        if center:
            cv.drawCentredString(W/2, y, ar(str(text)))
        else:
            cv.drawRightString(W - 2.3*cm, y, ar(str(text)))
        y -= size * 0.45 + gap

    def hline(w=0.8):
        cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(w)
        cv.line(2.2*cm, y, W-2.2*cm, y)

    # ── رأس الوثيقة ──────────────────────────────────────────────
    wr('الجمهورية الجزائرية الديمقراطية الشعبية', 12, center=True, gap=5)
    wr('وزارة الداخلية والجماعات المحلية والنقل', 11, center=True, gap=6)
    hline(); y -= 8
    wr(f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}', 10, center=True, gap=4)
    wr(f'رقم الهاتف : {markaz.get("رقم الهاتف","")}      البريد الإلكتروني : {markaz.get("البريد الإلكتروني","")}', 9, center=True, gap=4)
    hline(); y -= 14

    # ── العنوان ──────────────────────────────────────────────────
    cv.setFont(FONT_NAME, 15); cv.setFillColorRGB(0,0,0)
    cv.drawCentredString(W/2, y, ar('وصل تسديد حقوق التكوين'))
    y -= 8
    cv.setFont(FONT_NAME, 11)
    cv.drawCentredString(W/2, y, ar(f'رقم الوصل : {raqm_wasl}'))
    y -= 18

    # ── بيانات المتربص ───────────────────────────────────────────
    wr(f'تسلّمنا من السيد/السيدة : {mtrsh.get("الاسم واللقب","")}', 11, gap=8)
    wr(f'المسجل في تخصص : {takh}     —     الدورة رقم : {dawra}', 10, gap=14)

    # ── جدول المبالغ ─────────────────────────────────────────────
    # رسم جدول يدوي
    col_right = W - 2.3*cm
    col_mid   = W/2 + 1*cm
    row_h     = 0.75*cm
    tbl_top   = y
    tbl_rows  = [
        ('المبلغ الإجمالي لحقوق التكوين',  f'{total:,.2f} دج'),
        ('المبلغ المسدَّد سابقاً',          f'{paye_avant:,.2f} دج'),
        ('المبلغ المسدَّد حالياً',          f'{paye_now:,.2f} دج'),
        ('مجموع ما تم تسديده',             f'{total_paye:,.2f} دج'),
        ('باقي الدين',                      f'{reste:,.2f} دج'),
    ]

    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(0.5)
    for i, (label, val) in enumerate(tbl_rows):
        row_y = tbl_top - i * row_h
        # خلفية متناوبة
        if i % 2 == 0:
            cv.setFillColorRGB(0.95, 0.95, 0.95)
            cv.rect(2.2*cm, row_y - row_h + 3, W-4.4*cm, row_h, fill=1, stroke=0)
        # إبراز باقي الدين بالأحمر إن كان > 0
        txt_color = (0.7, 0, 0) if (i == 4 and reste > 0) else (0, 0, 0)
        cv.setFont(FONT_NAME, 10); cv.setFillColorRGB(*txt_color)
        cv.drawRightString(col_right, row_y - 5, ar(label))
        cv.drawString(2.5*cm, row_y - 5, ar(val))
        cv.setStrokeColorRGB(0.7,0.7,0.7); cv.setLineWidth(0.3)
        cv.line(2.2*cm, row_y - row_h + 3, W-2.2*cm, row_y - row_h + 3)

    # إطار الجدول
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(0.8)
    cv.rect(2.2*cm, tbl_top - len(tbl_rows)*row_h + 3, W-4.4*cm, len(tbl_rows)*row_h)
    y = tbl_top - len(tbl_rows)*row_h - 0.3*cm

    # ── طريقة الدفع ──────────────────────────────────────────────
    y -= 10
    wr('طريقة الدفع :    ☐ نقداً          ☐ حوالة بريدية          ☐ تحويل بنكي', 10, gap=16)

    # ── التوقيع ──────────────────────────────────────────────────
    wr(f'حُررت بـ {ville} في : {today}', 10, gap=25)
    wr('ختم وتوقيع ممثل مركز التكوين', 10, gap=0)

    # ── QR Code ──────────────────────────────────────────────────
    qr_text = (
        f"نوع الوثيقة: وصل تسديد\n"
        f"رقم الوصل: {raqm_wasl}\n"
        f"الاسم واللقب: {mtrsh.get('الاسم واللقب','')}\n"
        f"التخصص: {takh}\n"
        f"رقم الدورة: {dawra}\n"
        f"المبلغ الإجمالي: {total:,.2f} دج\n"
        f"المسدَّد سابقاً: {paye_avant:,.2f} دج\n"
        f"المسدَّد حالياً: {paye_now:,.2f} دج\n"
        f"باقي الدين: {reste:,.2f} دج\n"
        f"التاريخ: {today}"
    )
    _draw_qr_generic(cv, qr_text, 2*cm, 2*cm, size_cm=2.8)

    cv.save()
    print(f"وصل تسديد → {out}")
    return out


def print_aqd_takwin(mtrsh, raqm_aqd, output_dir=None):
    """
    طباعة عقد التكوين بين المركز والمتربص — صيغة A4.
    mtrsh    : قاموس بيانات المتربص
    raqm_aqd : رقم العقد التسلسلي
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm

    out_dir = output_dir or OUTPUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    _register_font()

    markaz    = load_markaz()
    W, H      = A4
    ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom       = str(mtrsh.get("الاسم واللقب","") or "متربص").replace(" ","_")
    raqm_safe = str(raqm_aqd).replace("/","_")
    out       = os.path.join(out_dir, f"عقد_تكوين_{raqm_safe}_{nom}_{ts}.pdf")

    takh   = _clean_takh(str(mtrsh.get("تخصص التكوين","") or ""))
    dawra  = str(mtrsh.get("رقم الدورة","") or "")
    total  = get_montant(takh)
    ville  = _clean_wilaya(markaz.get("مديرية النقل لولاية",""))
    today  = datetime.now().strftime("%Y/%m/%d")

    cv = rl_canvas.Canvas(out, pagesize=A4)
    cv.setPageSize(A4)

    # ── إطار مزدوج ───────────────────────────────────────────────
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(1.5)
    cv.rect(1.5*cm, 1.5*cm, W-3*cm, H-3*cm)
    cv.setLineWidth(0.5)
    cv.rect(1.9*cm, 1.9*cm, W-3.8*cm, H-3.8*cm)

    y = H - 2.6*cm

    def wr(text, size=10, center=False, gap=5, color=(0,0,0), left=False):
        nonlocal y
        cv.setFont(FONT_NAME, size); cv.setFillColorRGB(*color)
        if center:  cv.drawCentredString(W/2, y, ar(str(text)))
        elif left:  cv.drawString(2.3*cm, y, ar(str(text)))
        else:       cv.drawRightString(W-2.3*cm, y, ar(str(text)))
        y -= size * 0.45 + gap

    def hline(w=0.8):
        cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(w)
        cv.line(2.2*cm, y, W-2.2*cm, y)

    def banner(text):
        nonlocal y
        cv.setFillColorRGB(0.1, 0.2, 0.5)
        cv.rect(2.2*cm, y-3, W-4.4*cm, 17, fill=1, stroke=0)
        cv.setFont(FONT_NAME, 10); cv.setFillColorRGB(1,1,1)
        cv.drawCentredString(W/2, y+1, ar(text))
        y -= 22

    # ── رأس الوثيقة ──────────────────────────────────────────────
    wr('الجمهورية الجزائرية الديمقراطية الشعبية', 12, center=True, gap=5)
    wr('وزارة الداخلية والجماعات المحلية والنقل', 11, center=True, gap=6)
    hline(); y -= 8
    wr(f'مركز التكوين : {markaz.get("اسم مركز التكوين","")}', 10, center=True, gap=4)
    wr(f'رقم الاعتماد : {markaz.get("رقم الاعتماد","")}      رقم الهاتف : {markaz.get("رقم الهاتف","")}', 9, center=True, gap=4)
    hline(); y -= 12

    # ── العنوان ──────────────────────────────────────────────────
    cv.setFont(FONT_NAME, 15); cv.setFillColorRGB(0,0,0)
    cv.drawCentredString(W/2, y, ar('عقد تكوين'))
    y -= 8
    cv.setFont(FONT_NAME, 10)
    cv.drawCentredString(W/2, y, ar(f'رقم العقد : {raqm_aqd}'))
    y -= 16

    # ── طرفا العقد ───────────────────────────────────────────────
    wr('بين الطرفين الموقعين أدناه :', 10, gap=8)

    banner('الطرف الأول : مركز التكوين')
    wr(f'الاسم : {markaz.get("اسم مركز التكوين","")}', 10, gap=4)
    wr(f'ممثَّلاً بالسيد : {markaz.get("ممثل مركز التكوين","")}', 10, gap=4)
    wr(f'العنوان : {markaz.get("عنوان مركز التكوين","")}', 10, gap=4)
    wr(f'رقم الهاتف : {markaz.get("رقم الهاتف","")}', 10, gap=8)

    banner('الطرف الثاني : المتربص')
    wr(f'الاسم واللقب : {mtrsh.get("الاسم واللقب","")}', 10, gap=4)
    wr(f'تاريخ ومكان الازدياد : {mtrsh.get("تاريخ ومكان الازدياد","")}', 10, gap=4)
    wr(f'رقم التعريف الوطني : {mtrsh.get("رقم التعريف الوطني","")}', 10, gap=4)
    wr(f'العنوان : {mtrsh.get("العنوان","")}', 10, gap=4)
    wr(f'رقم الهاتف : {mtrsh.get("رقم هاتف المتربص","")}', 10, gap=10)

    hline(); y -= 10

    # ── المواد ───────────────────────────────────────────────────
    mawad = [
        ('المادة 1 — موضوع العقد',
         f'يلتزم مركز التكوين بتقديم تكوين مهني للمتربص في تخصص : {takh}  ضمن الدورة رقم : {dawra}'),
        ('المادة 2 — حقوق التكوين',
         f'حُدِّدت حقوق التكوين بمبلغ إجمالي قدره : {total:,.2f} دج  تُسدَّد حسب الاتفاق بين الطرفين.'),
        ('المادة 3 — التزامات مركز التكوين',
         'توفير الإطار البيداغوجي المؤهل — توفير الوسائل البيداغوجية — منح شهادة التسجيل — منح شهادة الكفاءة عند النجاح.'),
        ('المادة 4 — التزامات المتربص',
         'الانتظام في الحضور — احترام النظام الداخلي — تسديد الحقوق في الآجال المحددة — احترام الممتلكات.'),
        ('المادة 5 — فسخ العقد',
         'يُفسخ هذا العقد في حالة : عدم تسديد حقوق التكوين، الغياب المتكرر، أو الإخلال بالنظام الداخلي.'),
        ('المادة 6 — التحكيم',
         'في حالة النزاع يُرجع إلى الجهات القضائية المختصة.'),
    ]
    for titre, contenu in mawad:
        cv.setFont(FONT_NAME, 10); cv.setFillColorRGB(0.1, 0.2, 0.5)
        cv.drawRightString(W-2.3*cm, y, ar(titre))
        y -= 13
        cv.setFont(FONT_NAME, 9); cv.setFillColorRGB(0,0,0)
        cv.drawRightString(W-2.3*cm, y, ar(contenu))
        y -= 14

    hline(0.5); y -= 12

    # ── التوقيع ──────────────────────────────────────────────────
    wr(f'حُرِّر بـ {ville} في : {today}', 10, center=True, gap=16)
    cv.setFont(FONT_NAME, 10); cv.setFillColorRGB(0,0,0)
    cv.drawString(2.3*cm, y, ar('الطرف الثاني (المتربص)'))
    cv.drawRightString(W-2.3*cm, y, ar('الطرف الأول (ممثل المركز)'))
    y -= 30
    cv.drawString(2.3*cm, y, ar('التوقيع :'))
    cv.drawRightString(W-2.3*cm, y, ar('ختم وتوقيع :'))
    y -= 5
    # خطوط التوقيع
    cv.setStrokeColorRGB(0,0,0); cv.setLineWidth(0.5)
    cv.line(2.3*cm, y, 8*cm, y)
    cv.line(W-8*cm, y, W-2.3*cm, y)

    # ── QR Code ──────────────────────────────────────────────────
    qr_text = (
        f"نوع الوثيقة: عقد تكوين\n"
        f"رقم العقد: {raqm_aqd}\n"
        f"الاسم واللقب: {mtrsh.get('الاسم واللقب','')}\n"
        f"رقم التعريف الوطني: {mtrsh.get('رقم التعريف الوطني','')}\n"
        f"التخصص: {takh}\n"
        f"رقم الدورة: {dawra}\n"
        f"المبلغ الإجمالي: {total:,.2f} دج\n"
        f"المركز: {markaz.get('اسم مركز التكوين','')}\n"
        f"التاريخ: {today}"
    )
    _draw_qr_generic(cv, qr_text, 2*cm, 2*cm, size_cm=2.5)

    cv.save()
    print(f"عقد تكوين → {out}")
    return out


def open_pdf(path):
    if not path or not os.path.exists(path): return
    if   platform.system() == "Windows": os.startfile(path)
    elif platform.system() == "Darwin":  subprocess.Popen(["open",     path])
    else:                                subprocess.Popen(["xdg-open", path])


if __name__ == "__main__":
    path = print_dawra()
    if path: open_pdf(path)
