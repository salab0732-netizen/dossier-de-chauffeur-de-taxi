#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""حفظ الإعدادات وقراءتها من config.json بجانب البرنامج.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import json

from taxi.constants import CONFIG_FILE


def load_config() -> dict:
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_config(cfg: dict) -> bool:
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False
