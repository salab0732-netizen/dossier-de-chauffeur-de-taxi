#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""تطبيع نصّي مشترك — تُستعمل في مطابقة أسماء الأعمدة والقيم المحصورة."""
from __future__ import annotations

import re


def _norm(s: str) -> str:
    """يُسقط كل ما ليس حرفاً أو رقماً، ويُوحّد حالة الأحرف.

    تُطابَق بها أسماء أعمدة Excel والقيم المرجعية، فلا يُفرّق الفراغ
    ولا الشرطة ولا حالة الحرف بين «n-porte» و«N Porte» و«n_porte».
    """
    return re.sub(r"[^0-9a-z\u0621-\u064a]+", "", str(s or "").strip().lower())
