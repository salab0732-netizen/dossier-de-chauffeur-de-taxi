#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""مزوّد Anthropic Claude — مخطّط مُلزِم عبر tool use.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import json

import base64

from taxi.constants import CLAUDE_MODELS
from taxi.core.prompt import build_response_schema
from taxi.providers.base import _pkg_version
from taxi.providers.base import ProviderError, Cancelled, retry_call, sort_models


class ClaudeProvider:
    name = "Claude"

    def __init__(self, api_key: str, model: str, workspace_id: str = ""):
        self.api_key = api_key
        self.model = model
        self.workspace_id = (workspace_id or "").strip()
        self.on_debug = None

    def _client(self):
        try:
            import anthropic
        except ImportError:
            raise ProviderError("المكتبة غير مثبّتة — شغّل: pip install anthropic")
        kw = {"api_key": self.api_key}
        if self.workspace_id:
            kw["default_headers"] = {"anthropic-workspace-id": self.workspace_id}
        return anthropic.Anthropic(**kw)

    def check(self) -> str:
        try:
            self._client().messages.create(
                model=self.model, max_tokens=4,
                messages=[{"role": "user", "content": "hi"}])
        except Exception as e:
            if "workspace" in str(e).lower() and not self.workspace_id:
                raise ProviderError(
                    "المفتاح غير مرتبط بمساحة عمل.\n\n"
                    "الحل الأسرع: أنشئ مفتاحاً جديداً في console.anthropic.com\n"
                    "واختر مساحة عمل عند الإنشاء بدل Default.\n\n"
                    "أو الصق معرّف مساحة العمل في خانة «مساحة العمل» أدناه.")
            raise
        return "الاتصال ناجح"

    def list_models(self) -> list[str]:
        try:
            out = [m.id for m in self._client().models.list(limit=40)]
            return sort_models(out) if out else list(CLAUDE_MODELS)
        except Exception:
            return list(CLAUDE_MODELS)

    TOOL_NAME = "enregistrer_dossier_taxi"

    def extract(self, images: list[bytes], prompt: str):
        """يعيد dict مباشرة عبر الإخراج المُهيكل (tool use)."""
        client = self._client()
        content = [{
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg",
                       "data": base64.standard_b64encode(img).decode()},
        } for img in images]
        content.append({"type": "text", "text": prompt})

        tool = {
            "name": self.TOOL_NAME,
            "description": "تسجيل بيانات ملف رخصة استغلال سيارة أجرة المستخرجة من الوثائق.",
            "input_schema": build_response_schema("anthropic"),
        }
        msg = client.messages.create(
            model=self.model,
            max_tokens=16384,
            tools=[tool],
            tool_choice={"type": "tool", "name": self.TOOL_NAME},
            messages=[{"role": "user", "content": content}],
        )

        kinds = [getattr(b, "type", "?") for b in msg.content]
        stop = getattr(msg, "stop_reason", "?")
        if self.on_debug:
            self.on_debug(f"رد Claude: كتل={kinds or 'لا شيء'} · سبب التوقّف={stop}")

        for block in msg.content:
            if getattr(block, "type", "") == "tool_use":
                return dict(block.input)

        # احتياط: نموذج لم يستعمل الأداة
        texts = [b.text for b in msg.content if getattr(b, "type", "") == "text"]
        raw = "\n".join(texts).strip()
        if not raw:
            raise ProviderError(
                f"أعاد النموذج رداً فارغاً بلا بيانات.\n"
                f"الكتل المُستلمة: {kinds or 'لا شيء'} · سبب التوقّف: {stop}\n"
                f"إصدار مكتبة anthropic: {_pkg_version('anthropic')}\n\n"
                "إن كان الإصدار أقدم من 0.40 فحدّثه:  pip install -U anthropic")
        return raw
