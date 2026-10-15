#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""قراءة ردّ النموذج: تنظيف التعليقات والفواصل الزائدة ثم JSON.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import json
import re


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
