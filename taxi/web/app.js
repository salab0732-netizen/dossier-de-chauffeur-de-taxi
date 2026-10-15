/* ═══════════════════════════════════════════════════════════════════
   واجهة نظام إدخال بيانات سيارات الأجرة — ولاية البيض (32)

   الحالة كلّها في S؛ الخادم يبثّ التغيّرات عبر SSE فلا تُستجوَب
   نقاط النهاية في حلقة. التعديل في شاشة المراجعة يُرسَل فوراً
   ويعيد الخادم التنبيهات محسوبةً من جديد.
   ═══════════════════════════════════════════════════════════════════ */

'use strict';

const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

const S = {
  schema: null,
  providers: [],
  config: {},
  jobs: [],
  review: null,     // السجلّ المعروض
  page: 0,
  zoom: 1,
  group: 'operateur',
  browseFor: null,
  browsePath: '',
  totals: 0,        // مجموع السجلّات في Excel
};

const api = async (path, opts = {}) => {
  const r = await fetch(path, {
    headers: opts.body ? { 'Content-Type': 'application/json' } : {},
    ...opts,
  });
  if (!r.ok) throw new Error((await r.text().catch(() => '')) || `HTTP ${r.status}`);
  return r.status === 204 ? null : r.json();
};

const post = (path, body) =>
  api(path, { method: 'POST', body: body === undefined ? undefined : JSON.stringify(body) });

/* ── إشعارات ─────────────────────────────────────────────────── */

function toast(msg, kind = '') {
  const el = document.createElement('div');
  el.className = 'toast' + (kind ? ` is-${kind}` : '');
  el.textContent = msg;
  $('#toasts').append(el);
  setTimeout(() => {
    el.style.opacity = '0';
    setTimeout(() => el.remove(), 250);
  }, kind === 'err' ? 6500 : 3200);
}

/* ── السمة ───────────────────────────────────────────────────── */

const THEMES = { dark: 'فاتح', light: 'داكن' };

function setTheme(t) {
  document.documentElement.dataset.theme = t;
  $('#themeLabel').textContent = THEMES[t];
  try { localStorage.setItem('taxi-theme', t); } catch {}
}

$('#themeBtn').onclick = () =>
  setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark');

/* ── التنقّل ──────────────────────────────────────────────────── */

/* شريط الصدر ثابت — هويّة البرنامج وأرقامه الحيّة. عناوين الشاشات
   تحته، فلا يتكرّر العنوان مرّتين في الصفحة الواحدة. */

function show(view) {
  $$('.view').forEach(v => v.classList.toggle('is-active', v.dataset.view === view));
  $$('.nav-item').forEach(b => b.classList.toggle('is-active', b.dataset.view === view));
  if (view === 'records') { loadStats(); loadRecords(); }
  if (view === 'review')  renderPick();
  location.hash = view;
}

$$('.nav-item').forEach(b => b.onclick = () => show(b.dataset.view));

/* ── إحصاءات الصدر ───────────────────────────────────────────── */

function renderHero() {
  const c = {};
  for (const j of S.jobs) c[j.status] = (c[j.status] || 0) + 1;
  const rev = c['بانتظار المراجعة'] || 0;
  const run = c['قيد المعالجة'] || 0;
  const done = (c['مُدخَل'] || 0) + (c['مكرّر'] || 0);
  const n = S.schema ? S.schema.groups.reduce((a, g) => a + g.fields.length, 0) : 155;

  const bits = [
    { n: S.totals, l: 'سجلّاً في الملف' },
    { n, l: 'حقلاً لكل ملف' },
  ];
  if (run)  bits.push({ n: run,  l: 'قيد المعالجة' });
  if (rev)  bits.push({ n: rev,  l: 'بانتظار المراجعة', c: 'c-warn' });
  if (done) bits.push({ n: done, l: 'أُدخل اليوم', c: 'c-ok' });

  $('#heroStats').innerHTML = bits
    .map(b => `<div class="hstat ${b.c || ''}"><b>${b.n}</b><span>${b.l}</span></div>`)
    .join('');
}

/* ── حالة الملفات ────────────────────────────────────────────── */

const ST_CLASS = {
  'بانتظار': 'pending', 'قيد المعالجة': 'working', 'بانتظار المراجعة': 'review',
  'مُدخَل': 'saved', 'مكرّر': 'skipped', 'فشل': 'failed',
};

const NUMERIC = /nin|porte|permis|immatric|chasis|telephone|dossier|cnas|casnos|nif|nrc|siege|acte|certificat|decision|puissance|agrement|ligne|poids|charge/i;

/* ── الطابور ─────────────────────────────────────────────────── */

function renderQueue() {
  const body = $('#queueBody');
  body.textContent = '';
  $('#queueEmpty').classList.toggle('is-on', !S.jobs.length);

  for (const j of S.jobs) {
    const tr = document.createElement('tr');
    const pr = j.progress || {};
    const pct = Math.round((pr.frac || 0) * 100);
    const working = j.status === 'قيد المعالجة';

    tr.innerHTML = `
      <td class="mono">${j.id}</td>
      <td>
        <div class="cell-file">
          <b></b>
          <span></span>
        </div>
      </td>
      <td><span class="badge st-${ST_CLASS[j.status] || 'pending'}"></span></td>
      <td>${working
        ? `<div class="prog"><div class="prog-bar"><i style="width:${pct}%"></i></div><span></span></div>`
        : '<span style="color:var(--ink-3)">—</span>'}</td>
      <td>${warnCell(j)}</td>
      <td><div class="row-act"></div></td>`;

    tr.querySelector('.cell-file b').textContent = j.label || j.name;
    tr.querySelector('.cell-file span').textContent = j.name;
    tr.querySelector('.badge').textContent = j.status;
    if (working) tr.querySelector('.prog span').textContent = pr.stage || '';

    const act = tr.querySelector('.row-act');
    if (j.status === 'بانتظار المراجعة' || j.status === 'مُدخَل' || j.status === 'مكرّر') {
      act.append(mkBtn('مراجعة', 'btn-sm', () => openReview(j.id)));
    }
    if (j.status === 'فشل') {
      act.append(mkBtn('إعادة', 'btn-sm', async () => {
        await post(`/api/jobs/${j.id}/retry`); toast('أُعيد إلى الطابور');
      }));
    }
    if (working) {
      act.append(mkBtn('إلغاء', 'btn-sm btn-danger', () => post(`/api/jobs/${j.id}/cancel`)));
    }
    act.append(mkBtn('×', 'btn-icon btn-danger', async () => {
      await api(`/api/jobs/${j.id}`, { method: 'DELETE' });
      if (S.review?.id === j.id) { S.review = null; $('#split').hidden = true; }
    }));

    if (j.error) tr.title = j.error;
    body.append(tr);
  }

  renderCounters();
  renderHero();
  const rev = S.jobs.filter(j => j.status === 'بانتظار المراجعة').length;
  $('[data-count="queue"]').textContent = S.jobs.length || '';
  $('[data-count="review"]').textContent = rev || '';
}

/** خانة التنبيهات — «—» قبل المعالجة، لا «٠» التي تُفهَم أنه اجتاز الفحص. */
function warnCell(j) {
  if (j.status === 'بانتظار' || j.status === 'قيد المعالجة' || j.status === 'فشل')
    return '<span style="color:var(--ink-3)">—</span>';
  return j.warnings
    ? `<span class="wn-count has">${j.warnings}</span>`
    : '<span class="wn-count none">٠</span>';
}

function mkBtn(label, cls, fn) {
  const b = document.createElement('button');
  b.className = `btn ${cls}`;
  b.textContent = label;
  b.onclick = fn;
  return b;
}

function renderCounters() {
  const c = {};
  for (const j of S.jobs) c[j.status] = (c[j.status] || 0) + 1;
  const order = [
    ['بانتظار', ''], ['قيد المعالجة', 'is-run'], ['بانتظار المراجعة', 'is-warn'],
    ['مُدخَل', 'is-ok'], ['مكرّر', ''], ['فشل', 'is-err'],
  ];
  $('#counters').innerHTML = order
    .filter(([k]) => c[k])
    .map(([k, cls]) => `<div class="counter ${cls}"><i></i><b>${c[k]}</b>${k}</div>`)
    .join('');
}

/* ── الرفع والإضافة ──────────────────────────────────────────── */

async function upload(files) {
  const pdfs = [...files].filter(f => f.name.toLowerCase().endsWith('.pdf'));
  if (!pdfs.length) return toast('ملفات PDF فقط', 'warn');

  const fd = new FormData();
  pdfs.forEach(f => fd.append('files', f));
  try {
    const r = await fetch('/api/jobs/upload', { method: 'POST', body: fd });
    const d = await r.json();
    S.jobs = d.jobs; renderQueue();
    toast(`أُضيف ${d.added.length} ملفاً`, 'ok');
  } catch (e) {
    toast(`تعذّر الرفع: ${e.message}`, 'err');
  }
}

$('#filePick').onchange = e => { upload(e.target.files); e.target.value = ''; };

const drop = $('#drop');
['dragenter', 'dragover'].forEach(ev =>
  drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('is-over'); }));
['dragleave', 'drop'].forEach(ev =>
  drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('is-over'); }));
drop.addEventListener('drop', e => upload(e.dataTransfer.files));

$('#btnAddPath').onclick = async () => {
  const p = $('#pathInput').value.trim();
  if (!p) return;
  try {
    const d = await post('/api/jobs/path', { path: p });
    S.jobs = d.jobs; renderQueue();
    toast(`أُضيف ${d.added.length} ملفاً`, 'ok');
    $('#pathInput').value = '';
  } catch (e) { toast(e.message, 'err'); }
};

$('#pathInput').onkeydown = e => { if (e.key === 'Enter') $('#btnAddPath').click(); };

$('#btnClearDone').onclick = async () => {
  for (const j of S.jobs.filter(x => ['مُدخَل', 'مكرّر', 'فشل'].includes(x.status)))
    await api(`/api/jobs/${j.id}`, { method: 'DELETE' });
};

/* ── السجلّ ──────────────────────────────────────────────────── */

function addLog(msg, kind) {
  const box = $('#log');
  const d = document.createElement('div');
  d.innerHTML = `<span class="k-dim">${new Date().toLocaleTimeString('fr')}</span>`;
  const s = document.createElement('span');
  s.className = `k-${kind || 'info'}`;
  s.textContent = msg;
  d.append(s);
  box.append(d);
  while (box.children.length > 400) box.firstChild.remove();
  box.scrollTop = box.scrollHeight;
}

/* ── المراجعة ────────────────────────────────────────────────── */

function renderPick() {
  const list = S.jobs.filter(j => j.status !== 'بانتظار' && j.status !== 'قيد المعالجة');
  const box = $('#revPick');
  box.textContent = '';
  for (const j of list) {
    const saved = j.status === 'مُدخَل' || j.status === 'مكرّر';
    const tone = j.warnings ? 'w-warn' : saved ? 'w-saved' : 'w-clean';
    const b = document.createElement('button');
    b.className = `pick-card ${tone}` + (S.review?.id === j.id ? ' is-on' : '');
    b.innerHTML = '<span class="pick-ic"></span><span class="pick-tx"><b></b><span></span></span>';
    b.querySelector('.pick-ic').textContent = j.warnings ? j.warnings : saved ? '✓' : '٠';
    b.querySelector('.pick-tx b').textContent = j.label || j.name;
    b.querySelector('.pick-tx span').textContent =
      j.warnings ? 'تنبيهاً يحتاج نظرك' : saved ? j.status : 'اجتاز الفحص';
    b.onclick = () => openReview(j.id);
    box.append(b);
  }
  if (!list.length) $('#reviewSub').textContent = 'لا ملفات جاهزة للمراجعة بعد.';
}

async function openReview(jid) {
  show('review');
  try {
    S.review = await api(`/api/review/${jid}`);
  } catch (e) { return toast(e.message, 'err'); }
  S.page = 0; S.zoom = 1;
  $('#split').hidden = false;
  $('#reviewAct').hidden = false;
  $('#reviewSub').textContent = `${S.review.label} — ${S.review.name}`;
  renderPick();
  renderFields();
  renderWarns();
  loadPage();
}

function renderFields() {
  const box = $('#groups');
  box.textContent = '';

  for (const g of S.review.groups) {
    const nWarn  = g.fields.filter(f => f.state === 'warn').length;
    const nFill  = g.fields.filter(f => f.state !== 'empty').length;

    const d = document.createElement('details');
    d.className = `grp g-${g.key}`;
    d.open = g.active && (nWarn > 0 || nFill > 0);

    const sum = document.createElement('summary');
    sum.innerHTML = `
      <span class="grp-ic">${g.icon || ''}</span>
      <span>${g.title}</span>
      <span class="grp-meta">
        ${nWarn ? `<span class="pip pip-warn">${nWarn} مشكوك</span>` : ''}
        <span class="pip ${g.active ? 'pip-ok' : 'pip-off'}">${nFill}/${g.fields.length}</span>
      </span>`;
    d.append(sum);

    const body = document.createElement('div');
    body.className = 'grp-body';

    for (const f of g.fields) {
      const w = document.createElement('label');
      w.className = `f s-${f.state}` + (NUMERIC.test(f.key) ? ' is-num' : '');
      w.id = `f-${g.key}-${f.key}`;

      const lab = document.createElement('span');
      lab.className = 'f-label';
      lab.textContent = f.label;
      lab.title = f.hint || f.key;

      const inp = document.createElement('input');
      inp.type = 'text';
      inp.value = f.value ?? '';
      inp.placeholder = f.hint || '—';
      inp.dataset.grp = g.key;
      inp.dataset.key = f.key;
      inp.onchange = () => saveField(g.key, f.key, inp.value, w);

      w.append(lab, inp);
      body.append(w);
    }
    d.append(body);
    box.append(d);
  }
}

let saveTimer = null;
const pending = {};

function saveField(grp, key, val, el) {
  el.classList.add('is-edited');
  (pending[grp] ||= {})[key] = val;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(flushEdits, 450);
}

async function flushEdits() {
  if (!Object.keys(pending).length || !S.review) return;
  const edits = JSON.parse(JSON.stringify(pending));
  for (const k in pending) delete pending[k];
  try {
    const d = await post(`/api/review/${S.review.id}`, { edits, revalidate: true });
    S.review.warnings = d.warnings;
    renderWarns();
    repaintStates();
  } catch (e) { toast(e.message, 'err'); }
}

/** يُعيد تلوين الحقول حسب التنبيهات الجديدة دون إعادة بناء الحقول. */
function repaintStates() {
  const warns = S.review.warnings;
  for (const g of S.review.groups) {
    for (const f of g.fields) {
      const el = $(`#f-${g.key}-${f.key}`);
      if (!el) continue;
      const inp = el.querySelector('input');
      const empty = !inp.value;
      const hit = warns.some(w => f.label && w.includes(f.label));
      el.classList.remove('s-ok', 's-warn', 's-empty');
      el.classList.add(hit ? 's-warn' : empty ? 's-empty' : 's-ok');
    }
  }
}

function renderWarns() {
  const box = $('#warns');
  const pill = $('#warnPill');
  const w = S.review?.warnings || [];

  pill.hidden = false;
  pill.textContent = w.length ? `${w.length} تنبيهاً` : 'بلا تنبيهات ✓';
  pill.classList.toggle('is-clean', !w.length);

  if (!w.length) { box.hidden = true; return; }
  box.hidden = false;
  box.innerHTML = `<h3>
    <svg viewBox="0 0 20 20" style="width:15px;height:15px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round">
      <path d="M10 3 2.5 16.5h15L10 3Zm0 5v4m0 2.6v.1"/></svg>
    يحتاج نظرك</h3>`;
  for (const t of w) {
    const b = document.createElement('button');
    b.textContent = t;
    b.onclick = () => jumpTo(t);
    box.append(b);
  }
}

/** ينقل المؤشّر إلى الحقل الذي يذكره التنبيه بنصّ تسميته. */
function jumpTo(text) {
  for (const g of S.review.groups) {
    for (const f of g.fields) {
      if (!f.label || !text.includes(f.label)) continue;
      const el = $(`#f-${g.key}-${f.key}`);
      if (!el) continue;
      el.closest('details').open = true;
      el.scrollIntoView({ block: 'center', behavior: 'smooth' });
      el.classList.add('is-hit');
      setTimeout(() => el.classList.remove('is-hit'), 1500);
      el.querySelector('input').focus();
      return;
    }
  }
}

$('#btnDiag').onclick = async () => {
  try {
    const d = await post(`/api/review/${S.review.id}/diag`);
    if (d.ok) toast(`حُفظ التشخيص: ${d.pages} صفحة · ${d.empty} خانة فارغة`, 'ok');
    else toast(d.error || 'تعذّر الحفظ', 'err');
  } catch (e) { toast(e.message, 'err'); }
};

$('#btnRecheck').onclick = async () => {
  await flushEdits();
  const d = await post(`/api/review/${S.review.id}`, { edits: {}, revalidate: true });
  S.review.warnings = d.warnings;
  renderWarns(); repaintStates();
  toast(d.warnings.length ? `${d.warnings.length} تنبيهاً` : 'اجتاز الفحص', d.warnings.length ? 'warn' : 'ok');
};

$('#btnSaveXl').onclick = async () => {
  await flushEdits();
  if (S.review.warnings.length &&
      !confirmInline(`فيه ${S.review.warnings.length} تنبيهاً — الحفظ على أي حال؟`)) return;
  try {
    const d = await post(`/api/review/${S.review.id}/save`);
    if (d.ok) {
      toast(`حُفظ: ${(d.written || []).join(' · ') || 'لم يُكتب شيء'}`, 'ok');
      refreshJobs();
    } else toast(d.error || 'تعذّر الحفظ', 'err');
  } catch (e) { toast(e.message, 'err'); }
};

/** تأكيد داخل الصفحة — confirm() لا يعمل في كل البيئات. */
function confirmInline(msg) { return window.confirm(msg); }

/* ── الصورة الأصلية ──────────────────────────────────────────── */

function loadPage() {
  const n = S.review?.pages || 0;
  $('#pgLabel').textContent = n ? `${S.page + 1} / ${n}` : '—';
  $('#pgPrev').disabled = S.page <= 0;
  $('#pgNext').disabled = S.page >= n - 1;
  const img = $('#scanImg');
  if (!n) { img.removeAttribute('src'); return; }
  img.src = `/api/review/${S.review.id}/page/${S.page}`;
  applyZoom();
  loadOcr();
}

function applyZoom() {
  $('#scanImg').style.transform = `scale(${S.zoom})`;
  $('#zoomLabel').textContent = `${Math.round(S.zoom * 100)}%`;
}

/* ── نصّ المسح ───────────────────────────────────────────────── */

S.ocrOn = false;

async function loadOcr() {
  if (!S.ocrOn || !S.review) return;
  const meta = $('#ocrMeta'), box = $('#ocrText');
  box.textContent = 'جارٍ التحميل…';
  try {
    const d = await api(`/api/review/${S.review.id}/ocr/${S.page}`);
    box.textContent = d.available ? d.text : (d.note || '');
    meta.textContent = d.available ? `${d.chars} حرفاً · صفحة ${S.page + 1}` : '';
  } catch (e) {
    box.textContent = e.message;
    meta.textContent = '';
  }
}

$('#btnOcr').onclick = () => {
  S.ocrOn = !S.ocrOn;
  $('#btnOcr').classList.toggle('is-on', S.ocrOn);
  $('#ocrView').hidden = !S.ocrOn;
  $('#scanView').hidden = S.ocrOn;
  loadOcr();
};

$('#pgPrev').onclick = () => { if (S.page > 0) { S.page--; loadPage(); } };
$('#pgNext').onclick = () => { if (S.page < S.review.pages - 1) { S.page++; loadPage(); } };
$('#zoomIn').onclick  = () => { S.zoom = Math.min(S.zoom * 1.25, 6); applyZoom(); };
$('#zoomOut').onclick = () => { S.zoom = Math.max(S.zoom / 1.25, 0.5); applyZoom(); };

$('#scanView').addEventListener('wheel', e => {
  if (!e.ctrlKey) return;
  e.preventDefault();
  S.zoom = Math.min(Math.max(S.zoom * (e.deltaY < 0 ? 1.12 : 0.89), 0.5), 6);
  applyZoom();
}, { passive: false });

/* ── السجلّات ────────────────────────────────────────────────── */

/** الورقة ← مفتاح القسم، لتلوين البطاقة بلون قسمها. */
const SHEET_KEY = {
  'Operateur': 'operateur', 'Chauffeur': 'chauffeur',
  'Véhicule': 'vehicule', 'Licence Modjahid': 'licence_modjahid',
};

async function loadStats() {
  try {
    const d = await api('/api/stats');
    const sheets = d.sheets || [];
    S.totals = sheets.reduce((a, s) => a + s.count, 0);
    renderHero();
    $('#stats').innerHTML = sheets.map(s => `
      <div class="stat s-${SHEET_KEY[s.sheet] || 'operateur'}">
        <span class="stat-ic">${s.icon || ''}</span>
        <div><div class="stat-n">${s.count}</div><div class="stat-l">${s.title}</div></div>
      </div>`).join('');
  } catch { $('#stats').textContent = ''; }
}

function renderSheetTabs() {
  $('#sheetTabs').innerHTML = S.schema.groups
    .filter(g => g.key !== 'chauffeur_manaoub')
    .map(g => `<button class="sheet-tab t-${g.key}${g.key === S.group ? ' is-on' : ''}" data-g="${g.key}">${g.title}</button>`)
    .join('');
  $$('#sheetTabs .sheet-tab').forEach(b => b.onclick = () => {
    S.group = b.dataset.g; renderSheetTabs(); loadRecords();
  });
}

let recTimer = null;

async function loadRecords() {
  const q = $('#recQ').value.trim();
  let d;
  try { d = await api(`/api/records?group=${S.group}&q=${encodeURIComponent(q)}`); }
  catch (e) { return toast(e.message, 'err'); }

  if (d.error) {
    $('#recHead').textContent = ''; $('#recBody').textContent = '';
    $('#recEmpty').classList.add('is-on');
    $('#recEmpty').querySelector('strong').textContent = d.error;
    $('#recFoot').textContent = '';
    return;
  }

  $('#recHead').innerHTML = `<tr>${d.columns.map(c => `<th>${c.label}</th>`).join('')}</tr>`;
  $('#recBody').innerHTML = d.rows.map(r =>
    `<tr>${r.map((v, i) =>
      `<td class="${NUMERIC.test(d.columns[i].key) ? 'is-num' : ''}">${esc(v)}</td>`).join('')}</tr>`).join('');
  $('#recEmpty').classList.toggle('is-on', !d.rows.length);
  $('#recFoot').textContent = d.total
    ? `${d.shown} من ${d.total} سجلّاً · ورقة ${d.sheet}`
    : `ورقة ${d.sheet} — لا سجلّات`;
}

const esc = s => String(s ?? '').replace(/[&<>"]/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

$('#recQ').oninput = () => { clearTimeout(recTimer); recTimer = setTimeout(loadRecords, 260); };

/* ── الإعدادات ───────────────────────────────────────────────── */

function keyField(provider) {
  const map = { 'Gemini': 'gemini_key', 'Claude': 'claude_key', 'OCR.space': 'ocrspace_key' };
  return map[provider] || '';
}

const KEY_HINT = {
  'Gemini':    'طبقة مجانية سخيّة · aistudio.google.com/apikey',
  'Claude':    'مدفوع بالرصيد · console.anthropic.com',
  'OCR.space': 'مجاني: 500 طلب يومياً، 1 م.ب للصفحة · ocr.space/ocrapi/freekey',
};

function renderProviders() {
  const cur = S.config.provider || 'Gemini';
  $('#provGrid').innerHTML = S.providers.map(p => `
    <button class="prov-opt${p.name === cur ? ' is-on' : ''}" data-p="${p.name}">
      <b>${p.name}</b>
      <span>${p.needs_key ? 'يحتاج مفتاحاً' : 'بدون إنترنت'}</span>
    </button>`).join('');

  $$('#provGrid .prov-opt').forEach(b => b.onclick = () => {
    S.config.provider = b.dataset.p;
    const p = S.providers.find(x => x.name === b.dataset.p);
    S.config.model = p.models[0];
    renderProviders();
  });

  const p = S.providers.find(x => x.name === cur) || S.providers[0];
  $('#cfgModel').innerHTML = (p?.models || [])
    .map(m => `<option${m === S.config.model ? ' selected' : ''}>${m}</option>`).join('');

  const needs = p?.needs_key;
  $('#keyFld').hidden = !needs;
  $('#wsFld').hidden  = cur !== 'Claude';
  $('#keyHint').textContent = KEY_HINT[cur] || '';
  if (needs) $('#cfgKey').value = S.config[keyField(cur)] || '';

  $('#provName').textContent = `${cur} · ${S.config.model || '—'}`;
}

function fillSettings() {
  $('#cfgExcel').value = S.config.excel_path || '';
  $('#cfgWs').value    = S.config.workspace_id || '';
  $('#cfgOverwrite').checked = !!S.config.overwrite;
  $('#cfgAutoSave').checked  = !!S.config.auto_save;
  renderProviders();
}

$('#cfgModel').onchange = e => { S.config.model = e.target.value; renderProviders(); };

$('#btnSaveCfg').onclick = async () => {
  const cur = S.config.provider || 'Gemini';
  const body = {
    provider: cur,
    model: $('#cfgModel').value,
    excel_path: $('#cfgExcel').value.trim(),
    workspace_id: $('#cfgWs').value.trim(),
    overwrite: $('#cfgOverwrite').checked,
    auto_save: $('#cfgAutoSave').checked,
  };
  const kf = keyField(cur);
  if (kf) body[kf] = $('#cfgKey').value;
  try {
    await post('/api/settings', body);
    Object.assign(S.config, body);
    renderProviders();
    toast('حُفظت الإعدادات', 'ok');
  } catch (e) { toast(e.message, 'err'); }
};

$('#btnCheck').onclick = async () => {
  const out = $('#checkOut');
  const chip = $('#provChip');
  out.hidden = false;
  out.className = 'check-out';
  out.textContent = 'جارٍ الاختبار…';
  await $('#btnSaveCfg').onclick();
  try {
    const d = await post('/api/provider/check');
    out.classList.add(d.ok ? 'is-ok' : 'is-err');
    out.textContent = d.message;
    chip.className = 'prov-chip ' + (d.ok ? 'is-ok' : 'is-err');
  } catch (e) {
    out.classList.add('is-err');
    out.textContent = e.message;
    chip.className = 'prov-chip is-err';
  }
};

/* ── متصفّح الملفات ──────────────────────────────────────────── */

async function browse(path) {
  try {
    const d = await api(`/api/browse?path=${encodeURIComponent(path || '')}`);
    S.browsePath = d.path;
    $('#browsePath').textContent = d.path;
    const list = $('#browseList');
    list.textContent = '';

    if (d.parent) {
      list.append(mkBrowse('⬆', '..', () => browse(d.parent), false));
    }
    for (const e of d.entries) {
      list.append(mkBrowse(e.dir ? '📁' : '📄', e.name,
        () => e.dir ? browse(e.path) : pickFile(e.path), !e.dir));
    }
  } catch (e) { toast(e.message, 'err'); }
}

function mkBrowse(icon, name, fn, isFile) {
  const b = document.createElement('button');
  b.className = 'browse-item' + (isFile ? ' is-file' : '');
  b.innerHTML = '<i></i><span></span>';
  b.querySelector('i').textContent = icon;
  b.querySelector('span').textContent = name;
  b.onclick = fn;
  return b;
}

function pickFile(p) {
  if (S.browseFor === 'excel') $('#cfgExcel').value = p;
  closeBrowse();
}

const openBrowse = (what, start) => {
  S.browseFor = what;
  $('#browseModal').hidden = false;
  browse(start);
};
const closeBrowse = () => { $('#browseModal').hidden = true; };

$('#btnBrowseXl').onclick    = () => openBrowse('excel', $('#cfgExcel').value);
$('#browseClose').onclick = closeBrowse;
$('#browseCancel').onclick = closeBrowse;
$('#browsePick').onclick = () => pickFile(S.browsePath);
$('#browseModal').onclick = e => { if (e.target.id === 'browseModal') closeBrowse(); };

/* ── البثّ ───────────────────────────────────────────────────── */

function connect() {
  const es = new EventSource('/api/events');

  es.onmessage = ev => {
    let m;
    try { m = JSON.parse(ev.data); } catch { return; }

    if (m.type === 'queue') {
      S.jobs = m.jobs;
      renderQueue();
      if ($('.view[data-view="review"]').classList.contains('is-active')) renderPick();
    } else if (m.type === 'progress') {
      const j = S.jobs.find(x => x.id === m.job);
      if (j) { j.progress = { stage: m.stage, frac: m.frac }; renderQueue(); }
    } else if (m.type === 'log') {
      addLog(m.msg, m.kind);
      if (m.kind === 'err') toast(m.msg, 'err');
    }
  };

  es.onerror = () => { es.close(); setTimeout(connect, 2500); };
}

async function refreshJobs() {
  try { S.jobs = (await api('/api/jobs')).jobs; renderQueue(); } catch {}
}

/* ── الإقلاع ─────────────────────────────────────────────────── */

(async function init() {
  let saved = 'dark';
  try { saved = localStorage.getItem('taxi-theme') || 'dark'; } catch {}
  setTheme(saved === 'light' ? 'light' : 'dark');

  try {
    S.schema = await api('/api/schema');
    const st = await api('/api/state');
    S.config = st.config || {};
    S.providers = st.providers || [];
    S.jobs = st.jobs || [];
  } catch (e) {
    return toast(`تعذّر الاتصال بالخادم: ${e.message}`, 'err');
  }

  const n = S.schema.groups.reduce((a, g) => a + g.fields.length, 0);
  $('#appVer').textContent = `v${S.schema.app.version} · ${n} حقلاً`;

  fillSettings();
  renderSheetTabs();
  renderQueue();
  loadStats();          // لتملأ أرقام الصدر فوراً
  connect();

  const h = (location.hash || '').replace('#', '');
  show(['extract', 'review', 'records', 'settings'].includes(h) ? h : 'extract');

  // حفظ ما لم يُرسَل قبل إغلاق الصفحة
  window.addEventListener('beforeunload', () => {
    if (Object.keys(pending).length) flushEdits();
  });
})();
