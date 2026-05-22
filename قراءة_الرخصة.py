#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
وحدة قراءة رخصة السياقة الجزائرية باستخدام Claude Vision API
تستخرج كل البيانات تلقائياً وتملأ حقول المترشح
"""
import os, json, base64, re
from datetime import datetime

# ── استيراد Gemini مثل الكود الأصلي تماماً ───────────────────
try:
    from google import genai as _genai
    from PIL import Image as _PILImage
    import arabic_reshaper as _reshaper
    from bidi.algorithm import get_display as _get_display
    _GEMINI_AVAILABLE = True
except ImportError:
    _GEMINI_AVAILABLE = False

GEMINI_API_KEY = "AIzaSyC0Aom0fJY41PFFVB0h_foP0NMG7HiLbWg"

# مجلد حفظ صور الوثائق
DOCS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "صور_الوثائق")

def encode_image(path):
    """تحويل الصورة إلى base64"""
    with open(path, "rb") as f:
        return base64.standard_b64encode(f.read()).decode("utf-8")

def _get_api_key(service="anthropic"):
    """جلب مفتاح API"""
    if service == "gemini":
        # مفتاح Gemini مدمج مباشرة
        GEMINI_API_KEY = "AIzaSyC0Aom0fJY41PFFVB0h_foP0NMG7HiLbWg"
        key = os.environ.get("GEMINI_API_KEY", GEMINI_API_KEY)
        key_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gemini_key.txt")
        if not key and os.path.exists(key_file):
            with open(key_file, "r") as f: key = f.read().strip()
        return key
    else:  # anthropic
        key = os.environ.get("ANTHROPIC_API_KEY", "")
        key_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "api_key.txt")
        if not key and os.path.exists(key_file):
            with open(key_file, "r") as f: key = f.read().strip()
        return key


def extract_license_data_claude(image_path, prompt):
    """استخراج بيانات الرخصة باستخدام Claude API"""
    import urllib.request
    api_key = _get_api_key("anthropic")
    if not api_key:
        raise ValueError(
            "مفتاح Claude API غير موجود!\n"
            "أنشئ ملف 'api_key.txt' وضع فيه مفتاح Anthropic.\n"
            "https://console.anthropic.com/"
        )

    img_data  = encode_image(image_path)
    ext       = os.path.splitext(image_path)[1].lower()
    media_type= {".jpg":"image/jpeg",".jpeg":"image/jpeg",
                 ".png":"image/png",".webp":"image/webp"}.get(ext,"image/jpeg")

    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "messages": [{"role":"user","content":[
            {"type":"image","source":{"type":"base64","media_type":media_type,"data":img_data}},
            {"type":"text","text":prompt}
        ]}]
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={"Content-Type":"application/json",
                 "anthropic-version":"2023-06-01",
                 "x-api-key":api_key},
        method="POST"
    )
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read().decode("utf-8"))
    text = result["content"][0]["text"].strip()
    text = re.sub(r"```json|```","",text).strip()
    start = text.find("{"); end = text.rfind("}") + 1
    if start >= 0 and end > start: text = text[start:end]
    return json.loads(text)


def extract_license_data_gemini(image_path, prompt=None):
    """استخراج بيانات رخصة السياقة باستخدام Gemini API"""
    import urllib.request, base64 as _b64

    GEMINI_API_KEY = "AIzaSyC0Aom0fJY41PFFVB0h_foP0NMG7HiLbWg"

    # قراءة الصورة وتحويلها لـ base64
    ext = os.path.splitext(image_path)[1].lower()
    mime = {".jpg":"image/jpeg",".jpeg":"image/jpeg",
            ".png":"image/png",".bmp":"image/bmp"}.get(ext,"image/jpeg")

    with open(image_path, "rb") as f:
        img_b64 = _b64.standard_b64encode(f.read()).decode("utf-8")

    gemini_prompt = (
        "أنت مبرمج ومدقق بيانات خبير في رخص السياقة البيومترية الجزائرية. "
        "اقرأ الصورة بتركيز شديد واستخرج البيانات بالترتيب المكتوب بالظبط: "
        "1. اللقب: اكتب الكلمة العربية وبجانبها الكلمة الفرنسية (مثال: بلمأمون - BELMAMOUN). "
        "2. الاسم: اكتب الكلمة العربية وبجانبها الكلمة الفرنسية (مثال: جلول - DJELLOUL). "
        "3. تاريخ الميلاد (Date de naissance): استخرجه بدقة كما هو. "
        "4. مكان الميلاد: اكتبه بالعربية والفرنسية (مثال: البيض - EL BAYADH). "
        "5. رقم الرخصة (N° Permis): الرقم الذي يبدأ بحرف. "
        "6. الرقم الوطني: رقم التعريف الوطني الطويل (15-18 رقم). "
        "7. الجنس: M أو F. "
        "8. الأصناف الصالحة: مثل B. "
        "9. زمرة الدم. "
        "10. تاريخ الإصدار وتاريخ الانتهاء. "
        "نظم المخرجات في أسطر مرقمة واضحة."
    )

    payload = json.dumps({
        "contents": [{"parts": [
            {"inline_data": {"mime_type": mime, "data": img_b64}},
            {"text": gemini_prompt}
        ]}],
        "generationConfig": {"maxOutputTokens": 1000, "temperature": 0}
    }).encode("utf-8")

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"
    req = urllib.request.Request(url, data=payload,
                                  headers={"Content-Type": "application/json"},
                                  method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode("utf-8"))

    raw_text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
    return _parse_text_to_dict(raw_text)

    gemini_prompt = (
        "أنت مبرمج ومدقق بيانات خبير في رخص السياقة البيومترية الجزائرية. "
        "اقرأ الصورة بتركيز شديد واستخرج البيانات بالترتيب المكتوب بالظبط: "
        "1. اللقب: اكتب الكلمة العربية وبجانبها الكلمة الفرنسية المكتوبة في السطر رقم 1 (مثال: بلمأمون - BELMAMOUN). "
        "2. الاسم: اكتب الكلمة العربية وبجانبها الكلمة الفرنسية المكتوبة في السطر رقم 2 (مثال: جلول - DJELLOUL). "
        "3. تاريخ الميلاد (Date de naissance): استخرجه بدقة كما هو. "
        "4. مكان الميلاد (Lieu de naissance): اكتبه بالعربية والفرنسية (مثال: البيض - EL BAYADH). "
        "5. رقم الرخصة (N° Permis): اكتب الرقم التسلسلي الذي يبدأ بحرف L. "
        "6. الرقم الوطني: اكتب رقم التعريف الوطني الطويل. "
        "7. الجنس (Sexe): اكتب الجنس (ذكر - M) أو (أنثى - F). "
        "8. الأصناف الصالحة (Catégories): اكتب صنف السياقة الصالح فقط (مثل: B). "
        "نظم المخرجات في أسطر واضحة ومباشرة دون أي نصوص ترحيبية أو هوامش."
    )

    # نفس طريقة الكود الأصلي بالضبط
    client = _genai.Client(api_key=GEMINI_API_KEY)

    with _PILImage.open(image_path) as img:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[img, gemini_prompt]
        )

    raw_text = response.text.strip() if response.text else ""

    # fix_bilingual_text من الكود الأصلي
    def fix_bilingual_text(text):
        lines = text.split('\n')
        fixed = []
        for line in lines:
            if not line.strip():
                fixed.append("")
                continue
            reshaped = _reshaper.reshape(line)
            bidi_line = _get_display(reshaped)
            fixed.append(bidi_line)
        return '\n'.join(fixed)

    fix_bilingual_text(raw_text)  # للتحقق من صحة المعالجة
    return _parse_text_to_dict(raw_text)

    # نفس prompt الكود الأصلي بالضبط
    gemini_prompt = (
        "أنت مبرمج ومدقق بيانات خبير في رخص السياقة البيومترية الجزائرية. "
        "اقرأ الصورة بتركيز شديد واستخرج البيانات بالترتيب المكتوب بالظبط: "
        "1. اللقب: اكتب الكلمة العربية وبجانبها الكلمة الفرنسية المكتوبة في السطر رقم 1 (مثال: بلمأمون - BELMAMOUN). "
        "2. الاسم: اكتب الكلمة العربية وبجانبها الكلمة الفرنسية المكتوبة في السطر رقم 2 (مثال: جلول - DJELLOUL). "
        "3. تاريخ الميلاد (Date de naissance): استخرجه بدقة كما هو. "
        "4. مكان الميلاد (Lieu de naissance): اكتبه بالعربية والفرنسية (مثال: البيض - EL BAYADH). "
        "5. رقم الرخصة (N° Permis): اكتب الرقم التسلسلي الذي يبدأ بحرف L. "
        "6. الرقم الوطني: اكتب رقم التعريف الوطني الطويل. "
        "7. الجنس (Sexe): اكتب الجنس (ذكر - M) أو (أنثى - F). "
        "8. الأصناف الصالحة (Catégories): اكتب صنف السياقة الصالح فقط (مثل: B). "
        "نظم المخرجات في أسطر واضحة ومباشرة دون أي نصوص ترحيبية أو هوامش."
    )

    client = _genai.Client(api_key=GEMINI_API_KEY)

    # نفس طريقة الإرسال من الكود الأصلي
    with _Image.open(image_path) as img:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[img, gemini_prompt]
        )

    raw_text = response.text.strip() if response.text else ""

    # نفس دالة fix_bilingual_text من الكود الأصلي
    def fix_bilingual_text(text):
        lines = text.split('\n')
        fixed = []
        for line in lines:
            if not line.strip():
                fixed.append("")
                continue
            reshaped = _reshaper.reshape(line)
            bidi_line = _get_display(reshaped)
            fixed.append(bidi_line)
        return '\n'.join(fixed)

    fixed_text = fix_bilingual_text(raw_text)

    # تحويل النص لقاموس
    return _parse_text_to_dict(raw_text, fixed_text)


def _parse_text_to_dict(raw_text, fixed_text=""):
    """تحويل النص المستخرج من Gemini لقاموس البيانات"""
    import re as _re

    data = {
        "رقم_الرخصة": "", "اللقب": "", "الاسم": "",
        "اللقب_الفرنسي": "", "الاسم_الفرنسي": "",
        "تاريخ_الازدياد": "", "مكان_الازدياد": "",
        "رقم_التعريف_الوطني": "", "تاريخ_الإصدار": "",
        "تاريخ_الانتهاء": "", "مكان_الإصدار": "",
        "زمرة_الدم": "", "الجنس": "",
        "فئات_الرخصة": "", "العنوان": "",
        "ملاحظات": raw_text
    }

    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue

        # 1. اللقب
        if _re.match(r'^1[.)]\s*اللقب', line) or _re.match(r'^1[.)]\s*', line):
            parts = _re.split(r'\s*[-–]\s*', line, maxsplit=1)
            if len(parts) == 2:
                data["اللقب"] = _re.sub(r'^1[.)]\s*اللقب[:\s]*', '', parts[0]).strip()
                data["اللقب_الفرنسي"] = parts[1].strip()
            elif len(parts) == 1:
                data["اللقب"] = _re.sub(r'^1[.)]\s*(اللقب[:\s]*)?', '', parts[0]).strip()

        # 2. الاسم
        elif _re.match(r'^2[.)]\s*الاسم', line) or _re.match(r'^2[.)]\s*', line):
            parts = _re.split(r'\s*[-–]\s*', line, maxsplit=1)
            if len(parts) == 2:
                data["الاسم"] = _re.sub(r'^2[.)]\s*الاسم[:\s]*', '', parts[0]).strip()
                data["الاسم_الفرنسي"] = parts[1].strip()
            elif len(parts) == 1:
                data["الاسم"] = _re.sub(r'^2[.)]\s*(الاسم[:\s]*)?', '', parts[0]).strip()

        # 3. تاريخ الميلاد
        elif _re.match(r'^3[.)]\s*تاريخ', line):
            m = _re.search(r'\d{2}[./]\d{2}[./]\d{4}', line)
            if m:
                data["تاريخ_الازدياد"] = m.group().replace('.', '/')

        # 4. مكان الميلاد
        elif _re.match(r'^4[.)]\s*مكان', line):
            parts = _re.split(r'\s*[-–]\s*', line, maxsplit=1)
            loc = parts[-1].strip()
            loc = _re.sub(r'^4[.)]\s*مكان[^:]*[:\s]*', '', loc).strip()
            data["مكان_الازدياد"] = loc

        # 5. رقم الرخصة
        elif _re.match(r'^5[.)]\s*رقم الرخصة', line):
            m = _re.search(r'[A-Z]\d{5,}|\d{6,}', line)
            if m:
                data["رقم_الرخصة"] = m.group()

        # 6. الرقم الوطني
        elif _re.match(r'^6[.)]\s*الرقم الوطني', line):
            m = _re.search(r'\d{15,18}', line)
            if m:
                data["رقم_التعريف_الوطني"] = m.group()

        # 7. الجنس
        elif _re.match(r'^7[.)]\s*الجنس', line):
            if 'M' in line or 'ذكر' in line:
                data["الجنس"] = "M"
            elif 'F' in line or 'أنثى' in line:
                data["الجنس"] = "F"

        # 8. الأصناف
        elif _re.match(r'^8[.)]\s*الأصناف', line):
            m = _re.search(r'\b[ABCDE]\b', line.upper())
            if m:
                data["فئات_الرخصة"] = m.group()

    return data



def extract_license_data(image_path, service="auto"):
    """
    استخراج بيانات رخصة السياقة باستخدام Claude أو Gemini

    service: "claude" | "gemini" | "auto" (يجرب Claude أولاً ثم Gemini)
    """
    prompt = """أنت مبرمج ومدقق بيانات خبير في رخص السياقة البيومترية الجزائرية.
اقرأ الصورة بتركيز شديد واستخرج البيانات وأعد الإجابة بصيغة JSON فقط بدون أي نص إضافي:
{
  "رقم_الرخصة": "",
  "اللقب": "",
  "الاسم": "",
  "اللقب_الفرنسي": "",
  "الاسم_الفرنسي": "",
  "تاريخ_الازدياد": "",
  "مكان_الازدياد": "",
  "رقم_التعريف_الوطني": "",
  "تاريخ_الإصدار": "",
  "تاريخ_الانتهاء": "",
  "مكان_الإصدار": "",
  "زمرة_الدم": "",
  "الجنس": "",
  "فئات_الرخصة": "",
  "العنوان": "",
  "ملاحظات": ""
}

تعليمات:
- اللقب والاسم: بالعربية في حقول "اللقب" و"الاسم"، وبالفرنسية في "اللقب_الفرنسي" و"الاسم_الفرنسي"
- رقم التعريف الوطني: الرقم الطويل (15-18 رقم)
- تاريخ الازدياد والإصدار والانتهاء: بصيغة DD/MM/YYYY
- فئات الرخصة: مثل B, C, D أو 1, 2, 3
- الجنس: M أو F
- إذا لم تجد حقلاً اتركه فارغاً ""
"""

    errors = []

    if service in ("claude", "auto"):
        try:
            return extract_license_data_claude(image_path, prompt)
        except Exception as e:
            errors.append(f"Claude: {e}")
            if service == "claude":
                raise

    if service in ("gemini", "auto"):
        try:
            return extract_license_data_gemini(image_path, prompt)
        except Exception as e:
            errors.append(f"Gemini: {e}")
            if service == "gemini":
                raise

    raise Exception("فشل الاستخراج:\n" + "\n".join(errors))


def save_document_image(image_path, nin, doc_type="رخصة"):
    """
    حفظ صورة الوثيقة في مجلد المترشح
    doc_type: "رخصة" أو "شهادة"
    """
    import shutil
    os.makedirs(DOCS_DIR, exist_ok=True)
    ext = os.path.splitext(image_path)[1].lower() or ".jpg"
    dest = os.path.join(DOCS_DIR, f"{nin}_{doc_type}{ext}")
    shutil.copy2(image_path, dest)
    print(f"✅ صورة {doc_type} حُفظت → {dest}")
    return dest


def get_document_image(nin, doc_type="رخصة"):
    """جلب مسار صورة الوثيقة إن وجدت"""
    for ext in [".jpg", ".jpeg", ".png", ".webp"]:
        path = os.path.join(DOCS_DIR, f"{nin}_{doc_type}{ext}")
        if os.path.exists(path):
            return path
    return None


def format_date(d):
    """تحويل التاريخ لصيغة موحدة DD/MM/YYYY"""
    if not d: return ""
    d = str(d).strip()
    # DD.MM.YYYY → DD/MM/YYYY
    d = d.replace(".", "/").replace("-", "/")
    parts = d.split("/")
    if len(parts) == 3:
        p0, p1, p2 = parts
        # YYYY/MM/DD → DD/MM/YYYY
        if len(p0) == 4:
            return f"{p2}/{p1}/{p0}"
        return f"{p0}/{p1}/{p2}"
    return d


def map_to_mtrsh(data):
    """تحويل بيانات الرخصة إلى قاموس حقول المترشح"""
    nom_ar   = str(data.get("اللقب","") or "").strip()
    prenom_ar= str(data.get("الاسم","") or "").strip()
    nom_complet = f"{nom_ar} {prenom_ar}".strip() if nom_ar or prenom_ar else ""

    nom_fr   = str(data.get("اللقب_الفرنسي","") or data.get("NOM","") or "").strip()
    prenom_fr= str(data.get("الاسم_الفرنسي","") or data.get("PRENOM","") or "").strip()

    dob  = format_date(data.get("تاريخ_الازدياد",""))
    mob  = str(data.get("مكان_الازدياد","") or "").strip()
    dob_mob = f"{dob} - {mob}".strip(" -") if dob or mob else ""

    return {
        "الاسم واللقب":               nom_complet,
        "اللقب بالفرنسية":            nom_fr,
        "الاسم بالفرنسية":            prenom_fr,
        "تاريخ ومكان الازدياد":       dob_mob,
        "رقم التعريف الوطني":         str(data.get("رقم_التعريف_الوطني","") or "").strip(),
        "رقم رخصة السياقة":           str(data.get("رقم_الرخصة","") or "").strip(),
        "تاريخ صدور رخصة السياقة":    format_date(data.get("تاريخ_الإصدار","")),
        "تاريخ انتهاء رخصة السياقة":  format_date(data.get("تاريخ_الانتهاء","")),
        "مكان إصدار الرخصة":          str(data.get("مكان_الإصدار","") or "").strip(),
        "فئات الرخصة":                str(data.get("فئات_الرخصة","") or "").strip(),
        "زمرة الدم":                  str(data.get("زمرة_الدم","") or "").strip(),
        "الجنس":                      str(data.get("الجنس","") or "").strip(),
        "العنوان":                    str(data.get("العنوان","") or mob or "").strip(),
        "_raw": data
    }


if __name__ == "__main__":
    # اختبار
    import sys
    if len(sys.argv) > 1:
        img = sys.argv[1]
        print(f"تحليل: {img}")
        data = extract_license_data(img)
        print("البيانات المستخرجة:")
        for k, v in data.items():
            if v: print(f"  {k}: {v}")
        mapped = map_to_mtrsh(data)
        print("\nحقول المترشح:")
        for k, v in mapped.items():
            if v and k != "_raw": print(f"  {k}: {v}")


# ══════════════════════════════════════════════════════════════
# وحدة إرسال SMS عبر المودام
# ══════════════════════════════════════════════════════════════

import time

def detect_modem_port():
    """الكشف التلقائي عن منفذ المودام"""
    import serial.tools.list_ports
    ports = list(serial.tools.list_ports.comports())
    # نبحث عن مودام أو Huawei أو ZTE
    keywords = ["modem","huawei","zte","sierra","gsm","3g","4g","usb","mobile"]
    for p in ports:
        desc = (p.description or "").lower()
        if any(k in desc for k in keywords):
            return p.device
    # إذا لم نجد، نُرجع أول منفذ متاح
    return ports[0].device if ports else None

def send_sms(phone, message, com_port=None, baudrate=9600, timeout=10):
    """
    إرسال SMS عبر المودام

    المعاملات:
        phone    : رقم الهاتف (مثال: +213661234567 أو 0661234567)
        message  : نص الرسالة
        com_port : منفذ COM (مثال: COM3) — تلقائي إذا None
        baudrate : سرعة المنفذ (افتراضي 9600)
        timeout  : مهلة الانتظار بالثواني

    الإرجاع: (True, "تم الإرسال") أو (False, "رسالة الخطأ")
    """
    import serial

    # تنسيق رقم الهاتف
    phone = str(phone or "").strip().replace(" ","").replace("-","")
    if phone.startswith("0"):
        phone = "+213" + phone[1:]
    if not phone.startswith("+"):
        phone = "+213" + phone

    # الكشف عن المنفذ
    port = com_port or detect_modem_port()
    if not port:
        return False, "لم يتم العثور على مودام — تحقق من توصيل المودام"

    try:
        ser = serial.Serial(port, baudrate, timeout=timeout)
        time.sleep(0.5)

        def send_cmd(cmd, wait=1.0):
            ser.write((cmd + "\r").encode())
            time.sleep(wait)
            return ser.read(ser.in_waiting).decode("utf-8", errors="ignore")

        # تهيئة المودام
        r = send_cmd("AT")
        if "OK" not in r and "AT" not in r:
            ser.close()
            return False, f"المودام لا يستجيب على {port}"

        # وضع النص
        send_cmd("AT+CMGF=1")

        # ترميز Unicode للعربية
        send_cmd("AT+CSCS=\"UCS2\"")

        # تحويل الرسالة لـ UCS2
        msg_ucs2 = message.encode("utf-16-be").hex().upper()
        phone_ucs2 = phone.encode("utf-16-be").hex().upper()

        # إرسال
        send_cmd(f'AT+CMGS="{phone}"', wait=0.5)
        time.sleep(0.3)
        ser.write((message + "\x1A").encode("utf-8", errors="ignore"))
        time.sleep(3)
        response = ser.read(ser.in_waiting).decode("utf-8", errors="ignore")
        ser.close()

        if "+CMGS:" in response or "OK" in response:
            return True, f"✅ تم إرسال SMS إلى {phone}"
        elif "ERROR" in response:
            return False, f"خطأ في الإرسال: {response.strip()}"
        else:
            return True, f"✅ تم الإرسال (لم يتأكد)"

    except serial.SerialException as e:
        return False, f"خطأ في المنفذ {port}: {e}"
    except Exception as e:
        return False, f"خطأ: {e}"


def send_sms_bulk(contacts, message_template, com_port=None, progress_cb=None):
    """
    إرسال SMS جماعي

    المعاملات:
        contacts         : قائمة من (اسم, هاتف)
        message_template : نص الرسالة مع {الاسم} كمتغير
        com_port         : منفذ COM
        progress_cb      : دالة (i, total, nom, ok, msg) للتقدم

    الإرجاع: (عدد_الناجح, عدد_الفاشل, تفاصيل)
    """
    ok_count = 0; fail_count = 0; details = []

    for i, (nom, tel) in enumerate(contacts, 1):
        if not tel or not str(tel).strip():
            details.append((nom, False, "لا يوجد رقم هاتف"))
            fail_count += 1
            continue

        msg = message_template.replace("{الاسم}", str(nom))
        ok, result = send_sms(tel, msg, com_port)

        if ok: ok_count += 1
        else:  fail_count += 1
        details.append((nom, ok, result))

        if progress_cb:
            progress_cb(i, len(contacts), nom, ok, result)

        # انتظار بين الرسائل لتجنب ازدحام المودام
        time.sleep(2)

    return ok_count, fail_count, details
