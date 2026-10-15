#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
════════════════════════════════════════════════════════════════
  التعرّف الضوئي على النصوص عبر خدمة OCR.space
  عربي + إنجليزي — يعتمد على مكتبة requests وحدها
════════════════════════════════════════════════════════════════
  الاستعمال من الكود:
      from ocr_space import ocr_space_file, OcrSpaceError

      try:
          text = ocr_space_file("صورة.jpg")          # ← متغير نصي عادي
      except OcrSpaceError as e:
          print("فشل:", e)

  الاستعمال من سطر الأوامر:
      python ocr_space.py صورة.jpg             # يطبع النص المستخرج
      python ocr_space.py صورة.jpg --ara       # فرض العربية
      python ocr_space.py صورة.jpg --eng       # فرض الإنجليزية
      python ocr_space.py صورة.jpg --table     # جداول وبطاقات مقسّمة خانات

  المتطلبات:   pip install requests
  مفتاح مجاني: https://ocr.space/ocrapi/freekey
════════════════════════════════════════════════════════════════
"""
from __future__ import annotations

import os
import sys
import time

import requests

# ══════════════════════════════════════════════════════════════════
#  الإعدادات — عدّل هنا
# ══════════════════════════════════════════════════════════════════

# مفتاح API الخاص بك من ocr.space
OCR_SPACE_API_KEY = "ضع_مفتاحك_هنا"

# عنوان الخدمة (الخطة المجانية). المشتركون في PRO يضعون العنوان الذي وصلهم بالبريد.
OCR_SPACE_URL = "https://api.ocr.space/parse/image"

# اللغة: "auto" يكتشف العربية والإنجليزية معاً في الصورة نفسها،
#         أو "ara" للعربية وحدها، أو "eng" للإنجليزية وحدها
DEFAULT_LANGUAGE = "auto"

# المحرّك: 2 (الافتراضي، يدعم العربية والكشف التلقائي)
#          أو 3 (أدق في الجداول وأبطأ، وحصّته المجانية أصغر)
DEFAULT_ENGINE = 2

# أقصى حجم للملف: 1 ميغابايت في الخطة المجانية — ارفعه إن كانت خطتك مدفوعة
MAX_FILE_BYTES = 1024 * 1024

# مهلة انتظار الرد بالثواني، وعدد المحاولات عند انقطاع الاتصال أو انشغال الخادم
TIMEOUT_SECONDS = 60
MAX_RETRIES = 3

# الصيغ المقبولة
ALLOWED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".pdf"}

_PLACEHOLDER = "ضع_مفتاحك_هنا"


# ══════════════════════════════════════════════════════════════════
#  الأخطاء — كلها ترث من OcrSpaceError، فيكفي التقاطها وحدها
# ══════════════════════════════════════════════════════════════════

class OcrSpaceError(Exception):
    """خطأ عام من هذه الوحدة."""


class OcrSpaceImageError(OcrSpaceError):
    """الصورة نفسها: غير موجودة، فارغة، تالفة، صيغة غير مدعومة، أو حجم زائد."""


class OcrSpaceConnectionError(OcrSpaceError):
    """لا إنترنت، أو انتهت المهلة، أو الخادم مشغول — تفيد إعادة المحاولة لاحقاً."""


class OcrSpaceAuthError(OcrSpaceError):
    """المفتاح غير مُدخل أو مرفوض."""


class OcrSpaceQuotaError(OcrSpaceError):
    """تجاوز حصّة الطلبات (500 طلب يومياً في الخطة المجانية)."""


# بصمات بداية الملفات — للتأكد أن الملف صورة فعلاً قبل استهلاك طلب من الحصّة
_MAGIC = {
    b"\x89PNG\r\n\x1a\n": "PNG",
    b"\xff\xd8\xff": "JPG",
    b"GIF87a": "GIF",
    b"GIF89a": "GIF",
    b"BM": "BMP",
    b"II*\x00": "TIF",
    b"MM\x00*": "TIF",
    b"%PDF": "PDF",
}


def _detect_type(data: bytes) -> str | None:
    """نوع الملف من أول بايتاته (PNG/JPG/...) أو None إن لم يُعرف."""
    for sig, kind in _MAGIC.items():
        if data.startswith(sig):
            return kind
    return None


def _is_quota_message(text: str) -> bool:
    """رسالة الخدمة عند تجاوز الحصّة: «...upto maximum 500 number of times...»"""
    t = (text or "").lower()
    return "maximum" in t and "times" in t


# ══════════════════════════════════════════════════════════════════
#  الإرسال إلى الخدمة
# ══════════════════════════════════════════════════════════════════

def ocr_space_request(data: bytes, filename: str = "image.jpg", *,
                      language: str = DEFAULT_LANGUAGE,
                      engine: int = DEFAULT_ENGINE,
                      api_key: str | None = None,
                      is_table: bool = False,
                      scale: bool = True,
                      detect_orientation: bool = True,
                      overlay: bool = False) -> dict:
    """
    يرسل بايتات صورة إلى OCR.space ويعيد ردّ الخدمة كاملاً (dict) بعد التحقق
    من خلوّه من الأخطاء. أغلب البرامج تكتفي بـ ocr_space_bytes أو ocr_space_file.

    overlay=True يضيف مواقع الكلمات (TextOverlay) — مفيد لقراءة الجداول بالمواقع.
    """
    key = (api_key or OCR_SPACE_API_KEY or "").strip()
    if not key or key == _PLACEHOLDER:
        raise OcrSpaceAuthError("لم يُضبط مفتاح OCR.space — ضعه في OCR_SPACE_API_KEY "
                                "أعلى الملف (مفتاح مجاني: ocr.space/ocrapi/freekey).")

    # ── التحقق من الصورة محلياً قبل الإرسال ─────────────────────────
    if not data:
        raise OcrSpaceImageError("الصورة فارغة (0 بايت).")
    kind = _detect_type(data)
    if kind is None:
        raise OcrSpaceImageError("الملف ليس صورة صالحة (PNG/JPG/GIF/BMP/TIF/PDF) "
                                 "أو أنه تالف.")
    if len(data) > MAX_FILE_BYTES:
        raise OcrSpaceImageError(
            f"حجم الصورة {len(data) / 1048576:.2f} ميغابايت يتجاوز الحد المسموح "
            f"{MAX_FILE_BYTES / 1048576:.2f} ميغابايت — صغّر الصورة أو اخفض جودتها.")

    form = {
        "language": language,
        "OCREngine": str(engine),
        "isTable": str(is_table).lower(),
        "scale": str(scale).lower(),
        "detectOrientation": str(detect_orientation).lower(),
        "isOverlayRequired": str(overlay).lower(),
        "filetype": kind,                 # النوع صراحةً بدل الاعتماد على الامتداد
    }
    headers = {"apikey": key}             # المفتاح في الترويسة كما توصي وثائق الخدمة
    if not filename.isascii():            # أسماء الملفات العربية قد تُربك الخادم
        filename = "image." + kind.lower()

    # ── الإرسال، مع إعادة المحاولة للأعطال العابرة وحدها ───────────────
    last_err: OcrSpaceError | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.post(OCR_SPACE_URL, headers=headers, data=form,
                                 files={"file": (filename, data)},
                                 timeout=TIMEOUT_SECONDS)
        except requests.exceptions.Timeout as e:
            last_err = OcrSpaceConnectionError(
                f"انتهت مهلة الانتظار ({TIMEOUT_SECONDS} ثانية) دون رد من OCR.space.")
            last_err.__cause__ = e
        except requests.exceptions.ConnectionError as e:
            last_err = OcrSpaceConnectionError(
                "تعذّر الاتصال بـ OCR.space — تحقّق من الإنترنت أو من الجدار الناري.")
            last_err.__cause__ = e
        except requests.exceptions.RequestException as e:
            raise OcrSpaceConnectionError(f"خطأ في الطلب: {e}") from e
        else:
            body = resp.text or ""
            # تجاوز الحصّة قد يأتي برمز 403 أو 429 أو حتى 200 — نعرفه من نص الرسالة
            if resp.status_code == 429 or _is_quota_message(body):
                raise OcrSpaceQuotaError(
                    "تجاوزت حصّة الطلبات لـ OCR.space (500 طلب يومياً في الخطة "
                    "المجانية) — أعد المحاولة لاحقاً أو استعمل مفتاحاً مدفوعاً.")
            if resp.status_code in (401, 403):
                raise OcrSpaceAuthError(
                    f"مفتاح OCR.space مرفوض أو غير صالح (HTTP {resp.status_code}).")
            if resp.status_code >= 500:
                last_err = OcrSpaceConnectionError(
                    f"خادم OCR.space مشغول حالياً (HTTP {resp.status_code}).")
            elif resp.status_code != 200:
                raise OcrSpaceError(f"رد غير متوقع من OCR.space "
                                    f"(HTTP {resp.status_code}): {body[:200]}")
            else:
                return _check_response(resp)

        if attempt < MAX_RETRIES:
            time.sleep(2 * attempt)       # انتظار متزايد: 2 ثم 4 ثوانٍ
    raise last_err or OcrSpaceConnectionError("فشل غير معروف في الاتصال بـ OCR.space.")


def _check_response(resp: requests.Response) -> dict:
    """يحوّل الرد إلى dict ويرفع الخطأ المناسب إن أبلغت الخدمة عن فشل."""
    try:
        result = resp.json()
    except ValueError:
        raise OcrSpaceError(f"رد غير مفهوم من OCR.space: {resp.text[:200]}")

    if isinstance(result, str):           # أحياناً يعود الخطأ نصاً داخل JSON
        if _is_quota_message(result):
            raise OcrSpaceQuotaError("تجاوزت حصّة الطلبات لـ OCR.space.")
        raise OcrSpaceError(f"رد غير مفهوم من OCR.space: {result[:200]}")
    if not isinstance(result, dict):
        raise OcrSpaceError(f"رد غير مفهوم من OCR.space: {str(result)[:200]}")

    if result.get("IsErroredOnProcessing"):
        msg = result.get("ErrorMessage") or result.get("ErrorDetails") or "سبب غير مذكور"
        if isinstance(msg, list):
            msg = " | ".join(str(m) for m in msg)
        low = str(msg).lower()
        if "api key" in low or "apikey" in low:
            raise OcrSpaceAuthError(f"مفتاح OCR.space مرفوض: {msg}")
        if _is_quota_message(low):
            raise OcrSpaceQuotaError(f"تجاوزت حصّة الطلبات: {msg}")
        if any(w in low for w in ("file", "image", "format", "size", "corrupt", "pdf")):
            raise OcrSpaceImageError(f"رفضت OCR.space الصورة: {msg}")
        raise OcrSpaceError(f"فشلت المعالجة لدى OCR.space: {msg}")

    # OCRExitCode: 1 نجاح كامل · 2 نجاح جزئي · 3 فشل كل الصفحات · 4 خطأ فادح
    if result.get("OCRExitCode") in (3, 4) or not result.get("ParsedResults"):
        raise OcrSpaceImageError(
            f"لم تُقرأ الصورة (OCRExitCode={result.get('OCRExitCode')}).")
    return result


# ══════════════════════════════════════════════════════════════════
#  الدوال الجاهزة للاستعمال
# ══════════════════════════════════════════════════════════════════

def ocr_space_bytes(data: bytes, filename: str = "image.jpg", **options) -> str:
    """
    يستخرج النص من صورة في الذاكرة (bytes) ويعيده نصاً.
    يعيد "" إن كانت الصورة سليمة لكن بلا نص.
    الخيارات نفسها التي تقبلها ocr_space_request (language، engine، is_table...).
    """
    result = ocr_space_request(data, filename, **options)
    pages = []
    for page in result.get("ParsedResults", []):
        # FileParseExitCode = 1 يعني أن هذه الصفحة قُرئت بنجاح
        if page.get("FileParseExitCode") == 1:
            pages.append((page.get("ParsedText") or "").replace("\r\n", "\n").strip())
    return "\n\n".join(p for p in pages if p)


def ocr_space_file(path: str, **options) -> str:
    """
    يستخرج النص من صورة محفوظة على الجهاز ويعيده نصاً.

        text = ocr_space_file(r"C:\\صور\\بطاقة.jpg")
        text = ocr_space_file("ورقة.png", language="ara")
        text = ocr_space_file("جدول.jpg", is_table=True)
    """
    if not os.path.isfile(path):
        raise OcrSpaceImageError(f"الملف غير موجود: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise OcrSpaceImageError(f"صيغة غير مدعومة «{ext}» — المقبول: "
                                 + "، ".join(sorted(ALLOWED_EXTENSIONS)))
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError as e:
        raise OcrSpaceImageError(f"تعذّرت قراءة الملف: {e}") from e
    return ocr_space_bytes(data, os.path.basename(path), **options)


# ══════════════════════════════════════════════════════════════════
#  تشغيل مباشر من سطر الأوامر
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("الاستعمال: python ocr_space.py <مسار_الصورة> [--table] [--ara|--eng]")
        sys.exit(1)

    opts = {}
    if "--table" in sys.argv:
        opts["is_table"] = True
    if "--ara" in sys.argv:
        opts["language"] = "ara"
    if "--eng" in sys.argv:
        opts["language"] = "eng"

    try:
        extracted_text = ocr_space_file(sys.argv[1], **opts)
    except OcrSpaceError as err:
        print(f"❌ {err}")
        sys.exit(2)

    print(extracted_text if extracted_text else "(لم يُعثر على نص في الصورة)")
