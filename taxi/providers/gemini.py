#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""مزوّد Google Gemini — مخطّط مُلزِم عبر response_schema.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import json

from taxi.constants import GEMINI_MODELS
from taxi.core.prompt import build_response_schema
from taxi.providers.base import ProviderError, Cancelled, retry_call, sort_models


class GeminiProvider:
    name = "Gemini"

    def __init__(self, api_key: str, model: str):
        self.api_key = api_key
        self.model = model

    def check(self) -> str:
        try:
            from google import genai
        except ImportError:
            raise ProviderError("المكتبة غير مثبّتة — شغّل: pip install google-genai")
        client = genai.Client(api_key=self.api_key)
        names = []
        for m in client.models.list():
            n = getattr(m, "name", "") or ""
            if "generateContent" in (getattr(m, "supported_actions", None) or ["generateContent"]):
                names.append(n.replace("models/", ""))
        return f"الاتصال ناجح — {len(names)} نموذج متاح"

    def list_models(self) -> list[str]:
        from google import genai
        client = genai.Client(api_key=self.api_key)
        out = []
        for m in client.models.list():
            n = (getattr(m, "name", "") or "").replace("models/", "")
            if not n or any(w in n for w in ("embedding", "aqa", "imagen", "veo", "tts")):
                continue
            out.append(n)
        return sort_models(set(out))

    def extract(self, images: list[bytes], prompt: str) -> str:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=self.api_key)
        parts = [types.Part.from_bytes(data=img, mime_type="image/jpeg") for img in images]
        parts.append(types.Part.from_text(text=prompt))
        cfg = types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=build_response_schema("gemini"),   # بنية مضمونة
            temperature=0.0,
            max_output_tokens=16384,
        )
        try:
            resp = client.models.generate_content(model=self.model, contents=parts, config=cfg)
        except Exception as e:
            # بعض النماذج لا تدعم المخطط المُلزِم — نعيد المحاولة بدونه
            if "schema" not in str(e).lower() and "response_schema" not in str(e).lower():
                raise
            resp = client.models.generate_content(
                model=self.model, contents=parts,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0, max_output_tokens=16384),
            )
        txt = (resp.text or "").strip()
        if not txt:
            reason = "؟"
            try:
                reason = str(resp.candidates[0].finish_reason)
            except Exception:
                pass
            raise ProviderError(
                f"أعاد النموذج رداً فارغاً (سبب التوقّف: {reason}). "
                "جرّب نموذجاً آخر أو قلّل عدد الصفحات.")
        return txt
