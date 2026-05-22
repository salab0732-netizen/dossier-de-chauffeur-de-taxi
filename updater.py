#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
updater.py — نظام التحديث التلقائي عبر GitHub
المستودع: https://github.com/salab0732-netizen/courege-kafaa
"""
import os, sys, json, subprocess, shutil, hashlib, traceback
from datetime import datetime

GITHUB_USER   = "salab0732-netizen"
GITHUB_REPO   = "courege-kafaa"
GITHUB_BRANCH = "main"
VERSION_FILE  = "version.json"
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))

# الملفات التي تُحدَّث (البيانات لا تُمس أبداً)
UPDATE_FILES = [
    "نظام_شهادة_الكفاءة.py",
    "طباعة_الشهادات.py",
    "طباعة_وورد.py",
    "محاضر_الدورة.py",
    "قراءة_الرخصة.py",
    "قارئ_الرخصة.py",
    "database.py",
    "utils.py",
    "updater.py",
]

# الملفات المحمية — لا تُمس أبداً
PROTECTED = [
    "بيانات_مركز_التكوين.xlsx",
    "نظام_الكفاءة.db",
    "backup",
    "محاضر_PDF",
    "credentials.json",
]


def raw_url(filename):
    return f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{filename}"


def get_remote_version():
    """قراءة version.json من GitHub"""
    try:
        import urllib.request
        url = raw_url(VERSION_FILE)
        with urllib.request.urlopen(url, timeout=8) as r:
            return json.loads(r.read().decode())
    except Exception:
        return None


def get_local_version():
    """قراءة version.json المحلي"""
    path = os.path.join(BASE_DIR, VERSION_FILE)
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"version": "0.0.0", "date": ""}


def download_file(filename):
    """تنزيل ملف من GitHub"""
    import urllib.request
    url = raw_url(filename)
    dest = os.path.join(BASE_DIR, filename)
    tmp  = dest + ".tmp"
    try:
        with urllib.request.urlopen(url, timeout=30) as r:
            data = r.read()
        with open(tmp, "wb") as f:
            f.write(data)
        shutil.move(tmp, dest)
        return True
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        print(f"❌ فشل تنزيل {filename}: {e}")
        return False


def check_and_update(silent=False, parent_window=None):
    """
    التحقق من التحديثات وتطبيقها.
    silent=True: لا يعرض رسالة إذا لا يوجد تحديث
    parent_window: نافذة Tkinter للرسائل (اختياري)
    يُعيد True إذا حدث تحديث ويجب إعادة التشغيل
    """
    try:
        remote = get_remote_version()
        if not remote:
            return False  # لا إنترنت أو GitHub غير متاح

        local = get_local_version()

        rv = remote.get("version", "0.0.0")
        lv = local.get("version", "0.0.0")

        # مقارنة الإصدارات
        def ver_tuple(v):
            try: return tuple(int(x) for x in str(v).split("."))
            except: return (0, 0, 0)

        if ver_tuple(rv) <= ver_tuple(lv):
            if not silent:
                _show_msg(parent_window, "info",
                          "✅ البرنامج محدَّث",
                          f"النسخة الحالية {lv} هي الأحدث.")
            return False

        # يوجد تحديث
        notes = remote.get("notes", "")
        msg = (f"يتوفر تحديث جديد!\n\n"
               f"النسخة الحالية : {lv}\n"
               f"النسخة الجديدة : {rv}\n"
               f"التاريخ        : {remote.get('date','')}\n")
        if notes:
            msg += f"\nالتغييرات:\n{notes}"
        msg += "\n\nهل تريد التحديث الآن؟"

        if not _ask_msg(parent_window, "⬆️ تحديث متاح", msg):
            return False

        # تنزيل الملفات
        success = 0
        failed  = []
        files_to_update = remote.get("files", UPDATE_FILES)

        for fname in files_to_update:
            if any(fname.startswith(p) for p in PROTECTED):
                continue
            if download_file(fname):
                success += 1
            else:
                failed.append(fname)

        # تحديث version.json المحلي
        download_file(VERSION_FILE)

        if failed:
            _show_msg(parent_window, "warning",
                      "⚠️ تحديث جزئي",
                      f"تم تحديث {success} ملف.\n"
                      f"فشل تحديث: {', '.join(failed)}\n\n"
                      f"أعد تشغيل البرنامج.")
        else:
            _show_msg(parent_window, "info",
                      "✅ تم التحديث",
                      f"تم تحديث {success} ملف بنجاح!\n\n"
                      f"سيُعاد تشغيل البرنامج الآن.")

        return True  # يجب إعادة التشغيل

    except Exception as e:
        print(f"خطأ في التحديث: {e}")
        return False


def restart():
    """إعادة تشغيل البرنامج"""
    python = sys.executable
    os.execl(python, python, *sys.argv)


def _show_msg(parent, kind, title, message):
    try:
        import tkinter as tk
        from tkinter import messagebox
        if parent is None:
            root = tk.Tk(); root.withdraw()
            parent = root
        if kind == "info":
            messagebox.showinfo(title, message, parent=parent)
        elif kind == "warning":
            messagebox.showwarning(title, message, parent=parent)
        elif kind == "error":
            messagebox.showerror(title, message, parent=parent)
    except Exception:
        print(f"[{kind}] {title}: {message}")


def _ask_msg(parent, title, message):
    try:
        import tkinter as tk
        from tkinter import messagebox
        if parent is None:
            root = tk.Tk(); root.withdraw()
            parent = root
        return messagebox.askyesno(title, message, parent=parent)
    except Exception:
        return False


if __name__ == "__main__":
    updated = check_and_update(silent=False)
    if updated:
        restart()
