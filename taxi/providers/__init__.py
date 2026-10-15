# -*- coding: utf-8 -*-
"""مزوّدو الاستخراج: نموذجان سحابيان، ومحرّكا OCR."""
from taxi.providers.base import ProviderError, Cancelled, retry_call, sort_models
from taxi.providers.gemini import GeminiProvider
from taxi.providers.claude import ClaudeProvider
from taxi.providers.offline import OfflineProvider
from taxi.providers.ocrspace import OcrSpaceProvider

PROVIDERS = {
    "Gemini":    GeminiProvider,
    "Claude":    ClaudeProvider,
    "Offline":   OfflineProvider,
    "OCR.space": OcrSpaceProvider,
}

# المزوّدون الذين لا يحتاجون مفتاح API
NO_KEY = {"Offline"}

__all__ = ["PROVIDERS", "NO_KEY", "ProviderError", "Cancelled",
           "retry_call", "sort_models", "GeminiProvider", "ClaudeProvider",
           "OfflineProvider", "OcrSpaceProvider"]
