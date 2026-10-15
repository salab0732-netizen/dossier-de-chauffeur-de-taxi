#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ثوابت البرنامج — الاسم والإصدار والولاية ونماذج المزوّدين.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

from pathlib import Path


APP_NAME    = "نظام إدخال بيانات سيارات الأجرة"
APP_VERSION = "7.0"
WILAYA_ID   = 32
CONFIG_FILE = Path(__file__).resolve().parent.parent / "config.json"

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


OFFLINE_MODELS = ["tesseract-ara", "tesseract-ara+fra"]
OCRSPACE_MODELS = ["engine-2", "engine-3"]
