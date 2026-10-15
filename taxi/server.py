#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""خادم محلي — يخدم الواجهة ويُشغّل خطّ المعالجة.

يعمل على 127.0.0.1 وحده افتراضياً، فلا يُرى من الشبكة إلا بطلب صريح.
الحالة في الذاكرة: طابور الملفات وسجلّات المراجعة؛ ما يُحفظ فعلاً
يُكتب في Excel، وما يُعدّل يبقى في الذاكرة حتى يُحفظ.
"""
from __future__ import annotations

import asyncio
import json
import queue
import shutil
import tempfile
import threading
import traceback
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, UploadFile, File
from fastapi.responses import (FileResponse, HTMLResponse, JSONResponse,
                               Response, StreamingResponse)
from fastapi.staticfiles import StaticFiles

from taxi.config import load_config, save_config
from taxi.constants import (APP_NAME, APP_VERSION, CONFIG_FILE, CLAUDE_MODELS, GEMINI_MODELS,
                            OCRSPACE_MODELS, WILAYA_ID)
from taxi.core.enrich import NIN_FIELD, enrich, section_active
from taxi.core.normalize import normalize
from taxi.core.schema import SCHEMA, EXPECTED
from taxi.core.validate import validate
from taxi.excel import search as xsearch
from taxi.excel.writer import write_to_excel
from taxi.jobs import (ST_FAILED, ST_PENDING, ST_REVIEW, ST_SAVED, ST_SKIPPED,
                       ST_WORKING, Job, JobQueue)
from taxi.pdf import pdf_to_jpegs
from taxi.pipeline import run as run_pipeline
from taxi.providers import NO_KEY, PROVIDERS, Cancelled, ProviderError

WEB = Path(__file__).resolve().parent / "web"
UPLOADS = Path(tempfile.gettempdir()) / "taxi_uploads"
UPLOADS.mkdir(parents=True, exist_ok=True)

MODELS = {"Gemini": GEMINI_MODELS, "Claude": CLAUDE_MODELS,
          "OCR.space": OCRSPACE_MODELS}


# ══════════════════════════════════════════════════════════════════════
#  الحالة
# ══════════════════════════════════════════════════════════════════════

class State:
    """حالة الجلسة — طابور وسجلّات وبثّ الأحداث."""

    def __init__(self):
        self.cfg = load_config()
        self.queue = JobQueue(on_change=self._changed)
        self.jobs: dict[int, Job] = {}
        self.logs: dict[int, list] = {}
        self.pages: dict[int, list[bytes]] = {}
        self.ocr: dict[int, list[str]] = {}
        self.progress: dict[int, dict] = {}
        self.cancel: dict[int, threading.Event] = {}
        self._subs: list[queue.Queue] = []
        self._lock = threading.Lock()
        self._worker_stop = threading.Event()
        threading.Thread(target=self._worker, daemon=True).start()

    # ── بثّ الأحداث إلى كل واجهة مفتوحة ──
    def subscribe(self) -> queue.Queue:
        q: queue.Queue = queue.Queue(maxsize=256)
        with self._lock:
            self._subs.append(q)
        return q

    def unsubscribe(self, q):
        with self._lock:
            if q in self._subs:
                self._subs.remove(q)

    def emit(self, event: str, /, **payload):
        """يبثّ حدثاً إلى كل واجهة مفتوحة.

        `event` موضعيٌّ بحتاً (/) لأن الحمولة قد تحمل مفتاحاً بالاسم
        نفسه: `log` يبثّ حمولةً فيها `kind`، فلو كان اسم المعامل
        `kind` لتصادما وانهار كل استدعاء للسجلّ.
        """
        msg = json.dumps({"type": event, **payload}, ensure_ascii=False)
        with self._lock:
            subs = list(self._subs)
        for q in subs:
            try:
                q.put_nowait(msg)
            except queue.Full:
                pass

    def _changed(self):
        self.emit("queue", jobs=self.job_list())

    # ── الطابور ──
    def add(self, path, source="manual", note="") -> Job | None:
        job = self.queue.add(path, source, note)
        if job:
            self.jobs[job.id] = job
            self.logs[job.id] = []
            self.cancel[job.id] = threading.Event()
        return job

    def log(self, jid: int, msg: str, kind: str = "info"):
        entry = {"msg": msg, "kind": kind}
        self.logs.setdefault(jid, []).append(entry)
        self.emit("log", job=jid, **entry)

    def job_dict(self, j: Job) -> dict:
        return {
            "id": j.id, "name": j.name, "label": j.label(),
            "status": j.status, "source": j.source, "note": j.note,
            "error": j.error, "summary": j.summary,
            "warnings": len(j.warnings),
            "created": j.created.strftime("%H:%M"),
            "progress": self.progress.get(j.id, {}),
            "pages": len(self.pages.get(j.id, [])),
        }

    def job_list(self) -> list[dict]:
        return [self.job_dict(j) for j in self.queue.all()]

    # ── عامل المعالجة ──
    def _worker(self):
        while not self._worker_stop.is_set():
            job = self.queue.next_pending()
            if not job:
                self._worker_stop.wait(0.4)
                continue
            try:
                self._process(job)
            except Exception:
                job.status = ST_FAILED
                job.error = traceback.format_exc(limit=3)
                self.log(job.id, f"خطأ غير متوقّع: {job.error}", "err")
            self._changed()

    def _process(self, job: Job):
        self._changed()
        ev = self.cancel.setdefault(job.id, threading.Event())

        def prog(stage, frac):
            self.progress[job.id] = {"stage": stage, "frac": round(frac, 3)}
            self.emit("progress", job=job.id, stage=stage, frac=round(frac, 3))

        try:
            prov = self.provider()
        except ProviderError as e:
            job.status = ST_FAILED
            job.error = str(e)
            self.log(job.id, str(e), "err")
            return

        try:
            res = run_pipeline(
                str(job.path), prov, self.cfg.get("excel_path", ""),
                log=lambda m, k="info": self.log(job.id, m, k),
                progress=prog,
                should_stop=ev.is_set,
            )
        except Cancelled:
            job.status = ST_FAILED
            job.error = "أُلغيت بطلب المستخدم"
            self.log(job.id, "أُلغيت.", "warn")
            return
        except ProviderError as e:
            job.status = ST_FAILED
            job.error = str(e)
            self.log(job.id, str(e), "err")
            return

        job.data = res.data
        job.warnings = res.warnings
        self.ocr[job.id] = res.ocr
        try:
            self.pages[job.id] = pdf_to_jpegs(str(job.path), zoom=1.6)
        except Exception:
            self.pages[job.id] = []

        auto = bool(self.cfg.get("auto_save")) and res.clean
        if auto:
            self._save(job)
        else:
            job.status = ST_REVIEW

    def _save(self, job: Job) -> dict:
        xl = self.cfg.get("excel_path", "")
        if not xl or not Path(xl).exists():
            job.status = ST_REVIEW
            self.log(job.id, "ملف Excel غير محدّد أو غير موجود.", "err")
            return {"ok": False, "error": "ملف Excel غير محدّد أو غير موجود"}

        out = write_to_excel(xl, job.data, bool(self.cfg.get("overwrite")),
                             lambda m, k="info": self.log(job.id, m, k))
        wrote = out.get("written") or []
        skipped = out.get("skipped") or []
        job.summary = " · ".join(wrote) if wrote else "لم يُكتب شيء"
        job.status = ST_SKIPPED if (skipped and not wrote) else ST_SAVED
        self.log(job.id, f"حُفظ: {job.summary}", "ok")
        self._changed()
        return {"ok": True, **out}

    # ── المزوّد ──
    def _key_of(self, name: str) -> str:
        """مفتاح مزوّدٍ بعينه كما هو محفوظ في الإعدادات."""
        return (self.cfg.get(f"{name.lower().replace('.', '')}_key")
                or self.cfg.get(f"{name.lower()}_key") or "").strip()

    def provider(self):
        name = self.cfg.get("provider", "Gemini")
        cls = PROVIDERS.get(name)
        if not cls:
            # إعدادٌ يُسمّي مزوّداً لم يعد موجوداً — يُنقَل
            # إلى أوّل مزوّدٍ مفتاحُه محفوظ، فلا يُطالَب الموظّف بمفتاح
            # يملك غيره، ولا يقف البرنامج على إعدادٍ قديم.
            name = next((n for n in PROVIDERS if self._key_of(n)), "Gemini")
            cls = PROVIDERS[name]
            self.cfg["provider"] = name
            self.cfg["model"] = MODELS[name][0]
            save_config(self.cfg)
        model = self.cfg.get("model") or MODELS[name][0]
        if model not in MODELS[name]:
            model = MODELS[name][0]
        if name in NO_KEY:
            p = cls()
            p.model = model
            return p
        key = self._key_of(name)
        if name == "OCR.space":
            return cls(key, model)
        if not key:
            raise ProviderError(f"مفتاح {name} غير مُدخل — ضعه في الإعدادات.")
        if name == "Claude":
            return cls(key, model, self.cfg.get("workspace_id", ""))
        return cls(key, model)


ST = State()


# ══════════════════════════════════════════════════════════════════════
#  التطبيق
# ══════════════════════════════════════════════════════════════════════

app = FastAPI(title=APP_NAME, version=APP_VERSION, docs_url=None, redoc_url=None)


@app.get("/", response_class=HTMLResponse)
def index():
    return (WEB / "index.html").read_text(encoding="utf-8")


app.mount("/static", StaticFiles(directory=str(WEB)), name="static")


# ── المخطّط والحالة ───────────────────────────────────────────────────

@app.get("/api/schema")
def api_schema():
    """المخطّط كما تحتاجه الواجهة لرسم الحقول الـ155."""
    return {
        "app": {"name": APP_NAME, "version": APP_VERSION, "wilaya": WILAYA_ID},
        "groups": [
            {
                "key": k, "sheet": g["sheet"], "title": g["title"],
                "icon": g.get("icon", ""), "porte": g.get("porte"),
                "fields": [{"key": f[0], "label": f[1], "hint": f[2]}
                           for f in g["fields"]],
            }
            for k, g in SCHEMA.items()
        ],
    }


@app.get("/api/state")
def api_state():
    cfg = dict(ST.cfg)
    for k in list(cfg):
        if k.endswith("_key") or k == "mail_pass":
            cfg[k] = "••••" if cfg[k] else ""
    return {
        "config": cfg,
        "providers": [
            {"name": n, "models": MODELS[n], "needs_key": n not in NO_KEY}
            for n in PROVIDERS
        ],
        "jobs": ST.job_list(),
    }


@app.post("/api/settings")
async def api_settings(req: Request):
    body = await req.json()
    for k, v in body.items():
        if isinstance(v, str) and v == "••••":
            continue                       # لم يُغيّر المفتاح
        ST.cfg[k] = v
    save_config(ST.cfg)
    ST.emit("settings", config=True)
    return {"ok": True}


@app.post("/api/provider/check")
def api_provider_check():
    try:
        return {"ok": True, "message": ST.provider().check()}
    except ProviderError as e:
        return {"ok": False, "message": str(e)}
    except Exception as e:
        return {"ok": False, "message": f"{type(e).__name__}: {e}"}


# ── الطابور ──────────────────────────────────────────────────────────

@app.post("/api/jobs/upload")
async def api_upload(files: list[UploadFile] = File(...)):
    added = []
    for f in files:
        if not (f.filename or "").lower().endswith(".pdf"):
            continue
        dest = UPLOADS / f.filename
        i = 1
        while dest.exists():
            dest = UPLOADS / f"{Path(f.filename).stem}_{i}.pdf"
            i += 1
        with dest.open("wb") as fh:
            shutil.copyfileobj(f.file, fh)
        job = ST.add(dest, "manual", "")
        if job:
            added.append(ST.job_dict(job))
    return {"ok": True, "added": added, "jobs": ST.job_list()}


@app.post("/api/jobs/path")
async def api_add_path(req: Request):
    body = await req.json()
    p = Path(body.get("path", "")).expanduser()
    if not p.exists():
        raise HTTPException(404, "المسار غير موجود")
    paths = sorted(p.glob("*.pdf")) if p.is_dir() else [p]
    added = [ST.job_dict(j) for q in paths if (j := ST.add(q, "folder", ""))]
    return {"ok": True, "added": added, "jobs": ST.job_list()}


@app.get("/api/jobs")
def api_jobs():
    return {"jobs": ST.job_list()}


@app.post("/api/jobs/{jid}/cancel")
def api_cancel(jid: int):
    if jid in ST.cancel:
        ST.cancel[jid].set()
    return {"ok": True}


@app.post("/api/jobs/{jid}/retry")
def api_retry(jid: int):
    job = ST.jobs.get(jid)
    if not job:
        raise HTTPException(404, "لا يوجد")
    job.status = ST_PENDING
    job.error = ""
    job.warnings = []
    ST.cancel[jid] = threading.Event()
    ST._changed()
    return {"ok": True}


@app.delete("/api/jobs/{jid}")
def api_delete(jid: int):
    ST.queue.remove(jid)
    for d in (ST.jobs, ST.logs, ST.pages, ST.progress, ST.cancel):
        d.pop(jid, None)
    ST._changed()
    return {"ok": True}


@app.get("/api/events")
async def api_events(request: Request):
    """بثّ الأحداث — تقدّم، سجلّ، تغيّر الطابور."""
    q = ST.subscribe()

    async def gen():
        try:
            yield f"data: {json.dumps({'type': 'hello'})}\n\n"
            while True:
                if await request.is_disconnected():
                    break
                try:
                    msg = q.get_nowait()
                    yield f"data: {msg}\n\n"
                except queue.Empty:
                    await asyncio.sleep(0.25)
                    yield ": ping\n\n"
        finally:
            ST.unsubscribe(q)

    return StreamingResponse(gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache",
                                      "X-Accel-Buffering": "no"})


# ── المراجعة ─────────────────────────────────────────────────────────

def _field_state(grp: str, key: str, val, warnings: list[str]) -> str:
    """حالة الحقل كما تُعرَض: فارغ · مشكوك · مقروء."""
    if val in (None, ""):
        return "empty"
    label = dict((f[0], f[1]) for f in SCHEMA[grp]["fields"]).get(key, key)
    for w in warnings:
        if label and label in w:
            return "warn"
    return "ok"


@app.get("/api/review/{jid}")
def api_review(jid: int):
    job = ST.jobs.get(jid)
    if not job or not job.data:
        raise HTTPException(404, "لا توجد بيانات لهذا الملف")

    groups = []
    for k, g in SCHEMA.items():
        sec = job.data.get(k) or {}
        groups.append({
            "key": k, "title": g["title"], "icon": g.get("icon", ""),
            "active": section_active(job.data, k),
            "fields": [
                {"key": f[0], "label": f[1], "hint": f[2],
                 "value": sec.get(f[0], ""),
                 "state": _field_state(k, f[0], sec.get(f[0]), job.warnings)}
                for f in g["fields"]
            ],
        })

    return {
        "id": jid, "name": job.name, "label": job.label(),
        "status": job.status, "warnings": job.warnings,
        "logs": ST.logs.get(jid, []), "groups": groups,
        "pages": len(ST.pages.get(jid, [])),
    }


@app.post("/api/review/{jid}")
async def api_review_edit(jid: int, req: Request):
    """تعديل حقل واحد أو عدّة حقول، ثم إعادة الفحص فوراً."""
    job = ST.jobs.get(jid)
    if not job or not job.data:
        raise HTTPException(404, "لا توجد بيانات")

    body = await req.json()
    for grp, fields in (body.get("edits") or {}).items():
        if grp not in SCHEMA:
            continue
        job.data.setdefault(grp, {})
        for k, v in fields.items():
            job.data[grp][k] = v

    if body.get("revalidate", True):
        normalize(job.data)
        enrich(job.data)
    job.warnings = validate(job.data)
    return {"ok": True, "warnings": job.warnings}


@app.post("/api/review/{jid}/save")
def api_review_save(jid: int):
    job = ST.jobs.get(jid)
    if not job or not job.data:
        raise HTTPException(404, "لا توجد بيانات")
    return ST._save(job)


@app.get("/api/review/{jid}/ocr/{n}")
def api_ocr(jid: int, n: int):
    """نصّ OCR لصفحةٍ بعينها — ليُرى سببُ كل خانة فارغة.

    خانةٌ فارغة لها سببان لا ثالث لهما: إمّا أن المسح لم يقرأ القيمة،
    وإمّا أنها قُرئت والمستخرج لم يلتقطها. والفرق بينهما هو الفرق بين
    إصلاحٍ في محلّه وتخمينٍ في الهواء — ولا يُعرَف إلا بعرض النصّ.
    """
    pages = ST.ocr.get(jid) or []
    if not 0 <= n < len(pages):
        return {"text": "", "available": False,
                "note": "المزوّد الحالي لا يُعيد نصّاً خاماً — النماذج السحابية "
                        "تقرأ الصورة مباشرةً ولا تمرّ بنصّ وسيط."}
    return {"text": pages[n], "available": True, "chars": len(pages[n])}


@app.post("/api/review/{jid}/diag")
def api_diag(jid: int):
    """يكتب ملفّ تشخيص بجانب ملف الـPDF.

    يجمع في ملفٍّ واحد: نصّ المسح لكل صفحة، وتصنيفها، وما استُخرج
    منها، والتنبيهات. فيُقارَن ما قرأه المسح بما التقطه المستخرج
    حقلاً حقلاً — وهو ما لا يُستغنى عنه لإصلاح الأنماط، إذ لا يكفي
    أن نعرف أن الخانة فارغة، بل أين انقطع الطريق إليها.
    """
    job = ST.jobs.get(jid)
    if not job or not job.data:
        raise HTTPException(404, "لا توجد بيانات")

    ocr = ST.ocr.get(jid) or []
    try:
        from taxi.providers.regex_extract import _kind
        kinds = [sorted(_kind(p)) for p in ocr]
    except Exception:
        kinds = []

    empty = []
    for grp, fields in EXPECTED.items():
        if not section_active(job.data, grp):
            continue
        sec = job.data.get(grp) or {}
        labels = {f[0]: f[1] for f in SCHEMA[grp]["fields"]}
        for f in fields:
            if str(sec.get(f) or "").strip() == "":
                empty.append({"قسم": SCHEMA[grp]["title"], "حقل": f,
                              "تسمية": labels.get(f, f)})

    diag = {
        "الملف": job.name,
        "المزوّد": ST.cfg.get("provider"),
        "النموذج": ST.cfg.get("model"),
        "الإصدار": APP_VERSION,
        "عدد الصفحات": len(ocr),
        "التنبيهات": job.warnings,
        "خانات فارغة متوقَّعة": empty,
        "المستخرَج": job.data,
        "الصفحات": [{"رقم": i + 1, "تصنيف": kinds[i] if i < len(kinds) else [],
                     "حروف": len(t), "نصّ": t}
                    for i, t in enumerate(ocr)],
    }

    # يُكتب في مجلّد البرنامج لا بجانب الـPDF: الملف المسحوب يُنسَخ
    # إلى مجلّد Windows المؤقّت، فيضيع التشخيص هناك ولا يجده أحد.
    dest = CONFIG_FILE.parent / f"{job.path.stem}_تشخيص.json"
    try:
        dest.write_text(json.dumps(diag, ensure_ascii=False, indent=1),
                        encoding="utf-8")
    except Exception as e:
        return {"ok": False, "error": str(e)}
    return {"ok": True, "path": str(dest), "pages": len(ocr), "empty": len(empty)}


@app.get("/api/review/{jid}/page/{n}")
def api_page(jid: int, n: int):
    pages = ST.pages.get(jid) or []
    if not 0 <= n < len(pages):
        raise HTTPException(404, "لا توجد صفحة")
    return Response(pages[n], media_type="image/jpeg",
                    headers={"Cache-Control": "max-age=3600"})


# ── السجلّات ─────────────────────────────────────────────────────────

@app.get("/api/records")
def api_records(q: str = "", group: str = "operateur", limit: int = 200):
    xl = ST.cfg.get("excel_path", "")
    if not xl:
        return {"error": "ملف Excel غير محدّد", "rows": []}
    return xsearch.search(xl, q, group, limit)


@app.get("/api/stats")
def api_stats():
    xl = ST.cfg.get("excel_path", "")
    if not xl:
        return {"error": "ملف Excel غير محدّد", "sheets": []}
    return xsearch.stats(xl)


@app.get("/api/browse")
def api_browse(path: str = ""):
    """تصفّح مجلّدات الجهاز — لاختيار ملف Excel أو مجلّد المراقبة."""
    p = Path(path).expanduser() if path else Path.home()
    if not p.exists():
        p = Path.home()
    if p.is_file():
        p = p.parent
    try:
        entries = sorted(p.iterdir(), key=lambda x: (x.is_file(), x.name.lower()))
    except PermissionError:
        return {"path": str(p), "parent": str(p.parent), "entries": []}
    return {
        "path": str(p),
        "parent": str(p.parent) if p.parent != p else "",
        "entries": [
            {"name": e.name, "path": str(e), "dir": e.is_dir()}
            for e in entries
            if e.is_dir() or e.suffix.lower() in (".xlsx", ".xlsm", ".pdf")
        ][:500],
    }
