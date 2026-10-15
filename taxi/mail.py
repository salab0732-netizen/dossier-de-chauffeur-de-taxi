#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""قراءة مرفقات PDF من صندوق بريد IMAP إلى المجلّد المراقَب.

مُستخرَج من taxi_pro.py v5.7.1 — المنطق كما هو، مُجرَّب على ملفات حقيقية.
"""
from __future__ import annotations

import email
import email.header
import imaplib
import re
import threading
import time
import datetime
from pathlib import Path

from taxi.providers.base import ProviderError


IMAP_PRESETS = {
    "Gmail":          ("imap.gmail.com", 993),
    "Outlook/Hotmail": ("outlook.office365.com", 993),
    "Yahoo":          ("imap.mail.yahoo.com", 993),
    "آخر":            ("", 993),
}


class MailFetcher:
    """
    يتّصل بصندوق البريد عبر IMAP، ينزّل مرفقات PDF من الرسائل غير المقروءة،
    ويضعها في مجلّد المراقبة — فتدخل خطّ المعالجة نفسه.
    """

    def __init__(self, host, port, user, password, folder="INBOX",
                 subject_filter="", sender_filter="", mark_read=True):
        self.host, self.port = host, int(port or 993)
        self.user, self.password = user, password
        self.mailbox = folder or "INBOX"
        self.subject_filter = (subject_filter or "").strip()
        self.sender_filter = (sender_filter or "").strip()
        self.mark_read = mark_read

    def _connect(self):
        import imaplib
        if not self.host:
            raise ProviderError("خادم البريد غير محدّد.")
        if not self.user or not self.password:
            raise ProviderError("بريد المستخدم أو كلمة المرور غير مُدخلة.")
        try:
            m = imaplib.IMAP4_SSL(self.host, self.port, timeout=30)
        except Exception as e:
            raise ProviderError(f"تعذّر الاتصال بـ {self.host}:{self.port} — {e}")
        try:
            m.login(self.user, self.password)
        except Exception as e:
            msg = str(e)
            hint = ""
            if "gmail" in self.host.lower():
                hint = ("\n\nGmail يرفض كلمة المرور العادية. أنشئ «كلمة مرور التطبيقات» من:\n"
                        "myaccount.google.com ← الأمان ← التحقّق بخطوتين ← كلمات مرور التطبيقات")
            raise ProviderError(f"فشل تسجيل الدخول: {msg}{hint}")
        return m

    def check(self) -> str:
        m = self._connect()
        try:
            typ, data = m.select(self.mailbox, readonly=True)
            if typ != "OK":
                raise ProviderError(f"تعذّر فتح الصندوق «{self.mailbox}»")
            total = int(data[0]) if data and data[0] else 0
            typ, d = m.search(None, "UNSEEN")
            unseen = len(d[0].split()) if typ == "OK" and d and d[0] else 0
            return f"الاتصال ناجح — {total} رسالة، منها {unseen} غير مقروءة"
        finally:
            try:
                m.logout()
            except Exception:
                pass

    def _criteria(self):
        c = ["UNSEEN"]
        if self.sender_filter:
            c += ["FROM", f'"{self.sender_filter}"']
        if self.subject_filter:
            c += ["SUBJECT", f'"{self.subject_filter}"']
        return c

    def fetch_pdfs(self, dest_dir, log, limit=25) -> list:
        """ينزّل مرفقات PDF ويعيد قائمة (المسار، وصف الرسالة)."""
        import email
        from email.header import decode_header, make_header

        dest = Path(dest_dir)
        dest.mkdir(parents=True, exist_ok=True)
        out = []
        m = self._connect()
        try:
            if m.select(self.mailbox)[0] != "OK":
                raise ProviderError(f"تعذّر فتح الصندوق «{self.mailbox}»")
            typ, data = m.search(None, *self._criteria())
            ids = data[0].split() if typ == "OK" and data and data[0] else []
            if not ids:
                log("📭 لا رسائل جديدة مطابقة.", "dim")
                return out
            log(f"📬 {len(ids)} رسالة جديدة.", "info")

            for num in ids[-limit:]:
                typ, raw = m.fetch(num, "(RFC822)")
                if typ != "OK" or not raw or not raw[0]:
                    continue
                msg = email.message_from_bytes(raw[0][1])

                def hdr(h):
                    try:
                        return str(make_header(decode_header(msg.get(h, "") or "")))
                    except Exception:
                        return msg.get(h, "") or ""

                subj, frm = hdr("Subject"), hdr("From")
                found = 0
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    fname = part.get_filename()
                    if not fname:
                        continue
                    try:
                        fname = str(make_header(decode_header(fname)))
                    except Exception:
                        pass
                    if not fname.lower().endswith(".pdf"):
                        continue
                    payload = part.get_payload(decode=True)
                    if not payload:
                        continue
                    safe = "".join(ch for ch in fname
                                   if ch not in '\\/:*?"<>|').strip() or "mail.pdf"
                    target = dest / safe
                    i = 1
                    while target.exists():
                        target = dest / f"{Path(safe).stem}_{i}{Path(safe).suffix}"
                        i += 1
                    target.write_bytes(payload)
                    out.append((target, f"{frm} — {subj}"))
                    found += 1
                    log(f"   📎 {target.name}  ({frm})", "ok")

                if found and self.mark_read:
                    m.store(num, "+FLAGS", "\\Seen")
                elif not found:
                    log(f"   ✉️ بلا مرفق PDF: {subj[:50]}", "dim")
            return out
        finally:
            try:
                m.close()
            except Exception:
                pass
            try:
                m.logout()
            except Exception:
                pass


class MailPoller:
    """فحص دوري للبريد."""

    def __init__(self, make_fetcher, dest_dir, log, interval=300):
        self.make_fetcher = make_fetcher
        self.dest_dir = dest_dir
        self.log = log
        self.interval = max(60, int(interval))
        self._stop = threading.Event()
        self._thread = None
        self.last_run = None

    @property
    def running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.running:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        self.log(f"📧 فحص البريد كل {self.interval // 60} دقيقة.", "ok")

    def stop(self):
        self._stop.set()
        self.log("⏸ أُوقف فحص البريد.", "warn")

    def run_once(self) -> int:
        try:
            got = self.make_fetcher().fetch_pdfs(self.dest_dir, self.log)
            self.last_run = datetime.datetime.now()
            return len(got)
        except Exception as e:
            self.log(f"❌ البريد: {e}", "err")
            return 0

    def _loop(self):
        while not self._stop.is_set():
            self.run_once()
            self._stop.wait(self.interval)
