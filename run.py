#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
═══════════════════════════════════════════════════════════════════════
  نظام إدخال بيانات سيارات الأجرة  —  ولاية البيض (32)          v7.0
  يُشغّل الخادم المحلي ويفتح الواجهة.
═══════════════════════════════════════════════════════════════════════
  نافذة أصيلة (افتراضياً) :  python run.py
  في المتصفّح             :  python run.py --browser
  للشبكة المحلية          :  python run.py --host 0.0.0.0
═══════════════════════════════════════════════════════════════════════
"""
from __future__ import annotations

import argparse
import socket
import sys
import threading
import time
import webbrowser


def free_port(start: int = 8732, tries: int = 40) -> int:
    """أوّل منفذ شاغر — لا نُسقط الإقلاع لأن منفذاً واحداً مشغول."""
    for p in range(start, start + tries):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    raise SystemExit(f"لا منفذ شاغر بين {start} و{start + tries}.")


def wait_up(port: int, timeout: float = 25.0) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", port)) == 0:
                return True
        time.sleep(0.15)
    return False


def check_deps():
    missing = []
    for mod, pkg in [("fastapi", "fastapi"), ("uvicorn", "uvicorn[standard]"),
                     ("openpyxl", "openpyxl"), ("fitz", "pymupdf")]:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if missing:
        print("مكتبات ناقصة. ثبّتها بـ:\n")
        print(f"   pip install {' '.join(missing)}\n")
        raise SystemExit(1)


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("--host", default="127.0.0.1",
                    help="127.0.0.1 للجهاز وحده، 0.0.0.0 للشبكة المحلية")
    ap.add_argument("--port", type=int, default=0, help="0 = أوّل منفذ شاغر")
    ap.add_argument("--browser", action="store_true",
                    help="يفتح في المتصفّح بدل النافذة الأصيلة")
    ap.add_argument("--no-open", action="store_true", help="لا يفتح شيئاً")
    args = ap.parse_args()

    check_deps()

    import uvicorn
    from taxi.constants import APP_NAME, APP_VERSION

    port = args.port or free_port()
    url = f"http://127.0.0.1:{port}"

    print(f"\n  {APP_NAME}  v{APP_VERSION}")
    print(f"  {url}")
    if args.host == "0.0.0.0":
        try:
            lan = socket.gethostbyname(socket.gethostname())
            print(f"  الشبكة المحلية:  http://{lan}:{port}")
        except Exception:
            pass
    print("\n  للإيقاف:  Ctrl+C\n")

    cfg = uvicorn.Config("taxi.server:app", host=args.host, port=port,
                         log_level="warning", access_log=False)
    server = uvicorn.Server(cfg)

    if args.no_open:
        server.run()
        return

    # النافذة الأصيلة تحتاج الخادم قائماً، فيعمل في خيط
    threading.Thread(target=server.run, daemon=True).start()
    if not wait_up(port):
        raise SystemExit("لم يُقلع الخادم في الوقت المتوقّع.")

    if not args.browser:
        try:
            import webview                       # pywebview
            webview.create_window(f"{APP_NAME}  v{APP_VERSION}", url,
                                  width=1380, height=880, min_size=(900, 620))
            webview.start()
            return
        except ImportError:
            print("  pywebview غير مثبّت — يُفتح في المتصفّح.")
            print("  للنافذة الأصيلة:  pip install pywebview\n")

    webbrowser.open(url)
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n  أُوقف.\n")


if __name__ == "__main__":
    main()
