#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""أساس المزوّدين: الأخطاء وإعادة المحاولة وترتيب النماذج.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import random
import time


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
