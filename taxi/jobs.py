#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""الطابور ومراقبة المجلّد — كل PDF يصل يدخل خطّ المعالجة تلقائياً.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import datetime
import threading
from pathlib import Path


ST_PENDING  = "بانتظار"
ST_WORKING  = "قيد المعالجة"
ST_REVIEW   = "بانتظار المراجعة"
ST_SAVED    = "مُدخَل"
ST_SKIPPED  = "مكرّر"
ST_FAILED   = "فشل"

ST_COLOR = {
    ST_PENDING: "#64748b", ST_WORKING: "#1d4ed8", ST_REVIEW: "#b45309",
    ST_SAVED:   "#047857", ST_SKIPPED: "#7c3aed", ST_FAILED: "#b91c1c",
}

SRC_LABEL = {"manual": "يدوي", "folder": "مجلّد", "email": "بريد"}


class Job:
    """ملفٌّ واحد في رحلته عبر خطّ المعالجة."""
    _seq = 0

    def __init__(self, path, source="manual", note=""):
        Job._seq += 1
        self.id = Job._seq
        self.path = Path(path)
        self.source = source
        self.note = note
        self.status = ST_PENDING
        self.data: dict | None = None
        self.warnings: list[str] = []
        self.error = ""
        self.summary = ""
        self.created = datetime.datetime.now()

    @property
    def name(self):
        return self.path.name

    def label(self):
        if self.data:
            o = self.data.get("operateur") or {}
            nm = f"{o.get('nom_ar') or ''} {o.get('prenom_ar') or ''}".strip()
            pt = o.get("n-porte") or ""
            return f"{nm}  ·  باب {pt}" if nm or pt else self.name
        return self.name


class JobQueue:
    """طابور آمن للخيوط مع إشعار عند كل تغيّر."""

    def __init__(self, on_change=None):
        self._lock = threading.Lock()
        self._jobs: list[Job] = []
        self._seen: set[str] = set()
        self.on_change = on_change

    def _touch(self):
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass

    def add(self, path, source="manual", note="") -> "Job | None":
        key = str(Path(path).resolve()).lower()
        with self._lock:
            if key in self._seen:
                return None
            self._seen.add(key)
            job = Job(path, source, note)
            self._jobs.append(job)
        self._touch()
        return job

    def next_pending(self) -> "Job | None":
        with self._lock:
            for j in self._jobs:
                if j.status == ST_PENDING:
                    j.status = ST_WORKING
                    return j
        return None

    def update(self, job, **kw):
        with self._lock:
            for k, v in kw.items():
                setattr(job, k, v)
        self._touch()

    def all(self) -> list:
        with self._lock:
            return list(self._jobs)

    def get(self, jid) -> "Job | None":
        with self._lock:
            return next((j for j in self._jobs if j.id == jid), None)

    def counts(self) -> dict:
        with self._lock:
            c = {}
            for j in self._jobs:
                c[j.status] = c.get(j.status, 0) + 1
            return c

    def clear_done(self):
        with self._lock:
            self._jobs = [j for j in self._jobs
                          if j.status in (ST_PENDING, ST_WORKING, ST_REVIEW)]
        self._touch()

    def remove(self, jid) -> bool:
        """يُسقط ملفاً من الطابور، ويُنسى مساره فيُقبل مجدّداً لاحقاً."""
        with self._lock:
            for i, j in enumerate(self._jobs):
                if j.id == jid:
                    self._seen.discard(str(j.path.resolve()).lower())
                    del self._jobs[i]
                    break
            else:
                return False
        self._touch()
        return True
