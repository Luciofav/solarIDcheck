/* solarIDcheck — verifica seriali pannelli fotovoltaici
 * App statica: nessun server, i dati restano sul dispositivo (localStorage). */
'use strict';

const $ = (s) => document.querySelector(s);
const STORE_KEY = 'solaridcheck.session.v1';

const CFG = {
  interval: 110,      // ms tra un'analisi e la successiva
  minHits: 2,         // un codice è valido solo se letto in almeno 2 fotogrammi (anti lettura errata)
  pairWindow: 700,    // ms di attesa per trovare il secondo barcode
  candTTL: 600,       // ms dopo cui un candidato non più visto viene scartato
  cooldown: 2000,     // ms: un codice appena registrato viene ignorato finché resta inquadrato
  minLen: 4,          // lunghezza minima di un seriale valido
};
// Solo barcode lineari: restringere i formati velocizza e riduce le false letture.
const FORMATS = ['code_128', 'code_39', 'code_93', 'codabar', 'itf', 'ean_13', 'ean_8', 'upc_a', 'upc_e'];

/* ------------------------------------------------------------------ stato */

let session = null;      // { plant, startedAt, excel, scans: [], groupSeq }
let scanIndex = new Map();
let excelIndex = new Map();
let pendingExcel = null; // file caricato in fase di setup

function save() {
  try { localStorage.setItem(STORE_KEY, JSON.stringify(session)); } catch (e) { console.warn('save', e); }
}
function load() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch { return null; }
}
function clearStore() {
  try { localStorage.removeItem(STORE_KEY); } catch {}
}
function rebuildIndexes() {
  scanIndex = new Map(session.scans.map((s) => [s.serial, s]));
  excelIndex = new Map();
  for (const r of session.excel?.rows || []) if (r.serial && !excelIndex.has(r.serial)) excelIndex.set(r.serial, r);
}

const norm = (v) => String(v ?? '').replace(/\s+/g, '').trim().toUpperCase();

/* ------------------------------------------------------------------ viste */

function show(view) {
  for (const v of ['lock', 'setup', 'scan', 'report']) $('#view-' + v).hidden = v !== view;
  window.scrollTo(0, 0);
  if (view === 'scan') startCamera(); else stopCamera();
}

/* ------------------------------------------------------------------ Excel */

function cellText(cell) {
  if (!cell) return '';
  if (cell.t === 'n') {
    // Evita notazione scientifica (1,23E+15) per seriali numerici lunghi
    if (cell.w && !/e\+/i.test(cell.w)) return cell.w;
    return Number(cell.v).toLocaleString('fullwide', { useGrouping: false, maximumFractionDigits: 0 });
  }
  return String(cell.w ?? cell.v ?? '');
}

async function parseExcel(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellText: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws || !ws['!ref']) throw new Error('Il foglio è vuoto.');
  const range = XLSX.utils.decode_range(ws['!ref']);
  const grid = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) row.push(cellText(ws[XLSX.utils.encode_cell({ r, c })]).trim());
    grid.push({ excelRow: r + 1, cells: row });
  }
  // Prima riga non vuota = intestazione
  while (grid.length && grid[0].cells.every((c) => !c)) grid.shift();
  if (!grid.length) throw new Error('Il foglio è vuoto.');
  const header = grid.shift().cells;
  let iSer = header.findIndex((h) => /seri|s\/?n\b|matricol/i.test(h));
  let iMod = header.findIndex((h) => /model/i.test(h));
  if (iSer < 0) iSer = 0;
  if (iMod < 0) iMod = iSer === 0 ? 1 : 0;

  const rows = [];
  for (const g of grid) {
    const serial = norm(g.cells[iSer]);
    const model = (g.cells[iMod] || '').trim();
    if (!serial && !model) continue; // riga vuota
    rows.push({ serial, model, row: g.excelRow });
  }
  if (!rows.length) throw new Error('Nessuna riga dati trovata sotto l\'intestazione.');
  return {
    fileName: file.name,
    sheet: wb.SheetNames[0],
    colSerial: header[iSer] || 'colonna ' + (iSer + 1),
    colModel: header[iMod] || 'colonna ' + (iMod + 1),
    rows,
  };
}

function renderExcelPreview(box, ex, err) {
  box.hidden = false;
  if (err) { box.innerHTML = `<p class="err">${esc(err)}</p>`; return; }
  const models = new Set(ex.rows.map((r) => r.model).filter(Boolean));
  box.innerHTML = `
    <p><b>${ex.rows.length}</b> pannelli · <b>${models.size}</b> modelli · foglio “${esc(ex.sheet)}”</p>
    <table><tr><th>${esc(ex.colSerial)}</th><th>${esc(ex.colModel)}</th></tr>
    ${ex.rows.slice(0, 3).map((r) => `<tr><td>${esc(r.serial)}</td><td>${esc(r.model)}</td></tr>`).join('')}
    ${ex.rows.length > 3 ? '<tr><td colspan="2" class="muted">…</td></tr>' : ''}</table>`;
}

/* ------------------------------------------------------------------ setup */

function initSetup() {
  const saved = load();
  if (saved && saved.plant) {
    $('#resume-card').hidden = false;
    $('#resume-info').textContent =
      `${saved.plant} · ${saved.scans.length} scansioni · iniziata ${fmtDateTime(saved.startedAt)}` +
      (saved.excel ? ` · file ${saved.excel.fileName}` : ' · senza file impianto');
  }
  const plant = $('#plant');
  const refresh = () => { $('#btn-start').disabled = !plant.value.trim(); };
  plant.addEventListener('input', refresh);

  $('#excel').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    pendingExcel = null;
    if (!f) return;
    $('#excel-label').textContent = f.name;
    try {
      pendingExcel = await parseExcel(f);
      renderExcelPreview($('#excel-preview'), pendingExcel);
    } catch (err) {
      renderExcelPreview($('#excel-preview'), null, 'File non leggibile: ' + err.message);
    }
  });

  $('#btn-start').addEventListener('click', () => {
    const saved = load();
    if (saved && saved.scans?.length && !confirm(`C'è una sessione in corso (${saved.plant}, ${saved.scans.length} scansioni). Eliminarla e iniziarne una nuova?`)) return;
    session = {
      plant: plant.value.trim(),
      startedAt: Date.now(),
      excel: pendingExcel,
      scans: [],
      groupSeq: 0,
    };
    rebuildIndexes();
    save();
    unlockAudio();
    enterScan();
  });

  $('#btn-resume').addEventListener('click', () => {
    session = load();
    rebuildIndexes();
    unlockAudio();
    enterScan();
  });
  $('#btn-discard').addEventListener('click', () => {
    if (!confirm('Eliminare definitivamente la sessione salvata?')) return;
    clearStore();
    $('#resume-card').hidden = true;
  });
}

/* ------------------------------------------------------------------ scansione */

const scan = {
  stream: null, track: null, detector: null, timer: null,
  running: false, paused: false, busy: false,
  cand: new Map(),      // valore -> { hits, last }
  cooldown: new Map(),  // valore -> ultimo istante in cui è stato visto
  windowStart: null,
  torch: false, wakeLock: null,
};

function enterScan() {
  $('#scan-plant').textContent = session.plant;
  $('#scan-mode').textContent = session.excel
    ? `Confronto con ${session.excel.fileName}`
    : 'Senza file impianto: sola raccolta seriali';
  $('#c-expected-box').hidden = !session.excel;
  document.querySelector('.counters .ok').hidden = !session.excel;
  document.querySelector('.counters .bad').hidden = !session.excel;
  $('.counters').style.gridTemplateColumns = session.excel ? '' : '1fr';
  $('#toast').hidden = true;
  renderScanUI();
  show('scan');
}

async function getDetector() {
  if (scan.detector) return scan.detector;
  // 1) Rilevatore nativo (Chrome Android): il più veloce
  if ('BarcodeDetector' in window) {
    try {
      const sup = await window.BarcodeDetector.getSupportedFormats();
      const f = FORMATS.filter((x) => sup.includes(x));
      if (f.includes('code_128')) return (scan.detector = new window.BarcodeDetector({ formats: f }));
    } catch {}
  }
  // 2) ZXing in WebAssembly (iPhone, desktop): file .wasm servito dal repo, funziona offline
  const api = window.BarcodeDetectionAPI;
  api.prepareZXingModule({
    overrides: {
      locateFile: (path, prefix) => path.endsWith('.wasm') ? new URL('vendor/zxing_reader.wasm', location.href).href : prefix + path,
    },
  });
  return (scan.detector = new api.BarcodeDetector({ formats: FORMATS }));
}

async function startCamera() {
  if (scan.running) return;
  const msg = $('#cam-msg');
  msg.hidden = false;
  msg.textContent = 'Avvio fotocamera…';
  if (!navigator.mediaDevices?.getUserMedia) {
    msg.textContent = 'Fotocamera non disponibile: apri l\'app in HTTPS (es. GitHub Pages). Puoi inserire i seriali a mano.';
    return;
  }
  try {
    await getDetector();
    scan.stream = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } },
    });
  } catch (e) {
    msg.textContent = e.name === 'NotAllowedError'
      ? 'Permesso fotocamera negato. Abilitalo nelle impostazioni del browser.'
      : 'Impossibile avviare la fotocamera: ' + e.message;
    return;
  }
  if ($('#view-scan').hidden) { stopCamera(); return; } // l'utente è uscito nel frattempo
  const video = $('#video');
  video.srcObject = scan.stream;
  scan.track = scan.stream.getVideoTracks()[0];
  try { await video.play(); } catch {}

  const caps = scan.track.getCapabilities?.() || {};
  $('#btn-torch').hidden = !caps.torch;
  if (caps.focusMode?.includes('continuous')) {
    scan.track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(() => {});
  }
  try { scan.wakeLock = await navigator.wakeLock?.request('screen'); } catch {}

  msg.hidden = true;
  scan.running = true;
  scan.paused = false;
  setPauseUI();
  tick();
}

function stopCamera() {
  scan.running = false;
  clearTimeout(scan.timer);
  scan.stream?.getTracks().forEach((t) => t.stop());
  scan.stream = scan.track = null;
  scan.torch = false;
  $('#btn-torch').classList.remove('on');
  scan.cand.clear();
  scan.windowStart = null;
  scan.wakeLock?.release?.().catch(() => {});
  scan.wakeLock = null;
  clearOverlay();
}

async function tick() {
  if (!scan.running) return;
  const video = $('#video');
  if (!scan.paused && !scan.busy && video.readyState >= 2) {
    scan.busy = true;
    try {
      const codes = await scan.detector.detect(video);
      processFrame(codes, performance.now());
    } catch (e) { console.warn('detect', e); }
    scan.busy = false;
  }
  scan.timer = setTimeout(tick, CFG.interval);
}

/* Logica di cattura:
 * - un codice diventa "confermato" quando è letto in ≥ minHits fotogrammi;
 * - al primo confermato si apre una finestra di pairWindow ms in cui si cerca il secondo;
 * - se ne trova 2 li registra subito (come 2 pannelli distinti), altrimenti a fine finestra registra l'unico;
 * - i codici appena registrati restano in cooldown finché sono inquadrati, così non vengono contati due volte. */
function processFrame(codes, now) {
  const seen = [];
  for (const c of codes) {
    const v = norm(c.rawValue);
    if (v.length >= CFG.minLen && !seen.some((s) => s.v === v)) seen.push({ v, c });
  }

  for (const { v } of seen) if (scan.cooldown.has(v)) scan.cooldown.set(v, now);
  for (const [v, t] of scan.cooldown) if (now - t > CFG.cooldown) scan.cooldown.delete(v);

  for (const { v } of seen) {
    if (scan.cooldown.has(v)) continue;
    const k = scan.cand.get(v) || { hits: 0 };
    k.hits++; k.last = now;
    scan.cand.set(v, k);
  }
  for (const [v, k] of scan.cand) if (now - k.last > CFG.candTTL) scan.cand.delete(v);

  const confirmed = [...scan.cand].filter(([, k]) => k.hits >= CFG.minHits).map(([v]) => v);
  if (confirmed.length && scan.windowStart == null) scan.windowStart = now;
  if (scan.windowStart != null && (confirmed.length >= 2 || now - scan.windowStart >= CFG.pairWindow)) {
    if (confirmed.length) commitGroup(confirmed, now);
    scan.windowStart = null;
  }

  drawOverlay(seen);
}

function commitGroup(values, now) {
  session.groupSeq = (session.groupSeq || 0) + 1;
  const results = [];
  for (const v of values) {
    scan.cand.delete(v);
    scan.cooldown.set(v, now);
    results.push(addScan(v, 'camera', session.groupSeq, values.length));
  }
  feedback(results);
}

function addScan(serial, source, group, groupSize) {
  const ts = Date.now();
  const existing = scanIndex.get(serial);
  let status;
  if (existing) {
    existing.repeats.push(ts);
    status = 'dup';
  } else {
    const rec = { serial, ts, source, group, groupSize, repeats: [] };
    session.scans.push(rec);
    scanIndex.set(serial, rec);
    status = !session.excel || excelIndex.has(serial) ? 'ok' : 'unknown';
  }
  save();
  renderScanUI();
  return { serial, status, model: excelIndex.get(serial)?.model || '' };
}

function removeScan(serial) {
  session.scans = session.scans.filter((s) => s.serial !== serial);
  scanIndex.delete(serial);
  save();
  renderScanUI();
}

/* ------------------------------------------------------------------ feedback */

let audioCtx = null;
function unlockAudio() {
  try {
    audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
    audioCtx.resume();
  } catch {}
}
function beep(freq, dur, times = 1) {
  if (!audioCtx) return;
  for (let i = 0; i < times; i++) {
    const t0 = audioCtx.currentTime + i * (dur + 0.06);
    const o = audioCtx.createOscillator();
    const g = audioCtx.createGain();
    o.frequency.value = freq;
    o.type = 'square';
    g.gain.setValueAtTime(0.12, t0);
    g.gain.exponentialRampToValueAtTime(0.001, t0 + dur);
    o.connect(g).connect(audioCtx.destination);
    o.start(t0); o.stop(t0 + dur);
  }
}

function feedback(results) {
  const worst = results.some((r) => r.status === 'dup') ? 'dup'
    : results.some((r) => r.status === 'unknown') ? 'unknown' : 'ok';
  const label = { ok: 'OK', unknown: 'NON IN ELENCO', dup: 'GIÀ SCANSIONATO' };
  const toast = $('#toast');
  toast.className = 'toast ' + worst;
  toast.innerHTML = (results.length > 1 ? `<small>2 barcode letti → 2 pannelli</small>` : '') +
    results.map((r) => `${esc(r.serial)} — ${label[r.status]}${r.model ? `<small>${esc(r.model)}</small>` : ''}`).join('<br>');
  toast.hidden = false;

  if (worst === 'ok') { beep(1400, 0.08, results.length); navigator.vibrate?.(60); }
  else if (worst === 'unknown') { beep(700, 0.12, 2); navigator.vibrate?.([80, 60, 80]); }
  else { beep(300, 0.35); navigator.vibrate?.([250, 80, 250]); }
}

/* ------------------------------------------------------------------ overlay */

function clearOverlay() {
  const cv = $('#overlay');
  cv.getContext('2d').clearRect(0, 0, cv.width, cv.height);
}
function drawOverlay(seen) {
  const cv = $('#overlay');
  const video = $('#video');
  const dpr = window.devicePixelRatio || 1;
  const cw = cv.clientWidth, ch = cv.clientHeight;
  if (cv.width !== cw * dpr || cv.height !== ch * dpr) { cv.width = cw * dpr; cv.height = ch * dpr; }
  const ctx = cv.getContext('2d');
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cw, ch);
  const vw = video.videoWidth, vh = video.videoHeight;
  if (!vw || !vh) return;
  const s = Math.max(cw / vw, ch / vh); // object-fit: cover
  const ox = (cw - vw * s) / 2, oy = (ch - vh * s) / 2;
  ctx.lineWidth = 3;
  for (const { v, c } of seen) {
    const done = scan.cooldown.has(v);
    const color = !done ? '#ffd54a' : scanIndex.get(v)?.repeats.length ? '#ff5252' : '#4cd97b';
    ctx.strokeStyle = color;
    const pts = c.cornerPoints?.length ? c.cornerPoints : rectPoints(c.boundingBox);
    ctx.beginPath();
    pts.forEach((p, i) => (i ? ctx.lineTo : ctx.moveTo).call(ctx, ox + p.x * s, oy + p.y * s));
    ctx.closePath();
    ctx.stroke();
  }
}
function rectPoints(b) {
  return [{ x: b.x, y: b.y }, { x: b.x + b.width, y: b.y }, { x: b.x + b.width, y: b.y + b.height }, { x: b.x, y: b.y + b.height }];
}

/* ------------------------------------------------------------------ UI scansione */

function renderScanUI() {
  const total = session.scans.length;
  $('#c-scanned').textContent = total;
  if (session.excel) {
    const expected = excelIndex.size;
    const ok = session.scans.filter((s) => excelIndex.has(s.serial)).length;
    $('#c-expected').textContent = expected;
    $('#c-ok').textContent = ok;
    $('#c-unknown').textContent = total - ok;
    $('#progress').hidden = false;
    $('#progress i').style.width = (expected ? Math.min(100, (ok / expected) * 100) : 0) + '%';
  } else {
    $('#progress').hidden = true;
  }

  const list = $('#scan-list');
  const recent = session.scans.slice(-60).reverse();
  $('#list-count').textContent = total > 60 ? `ultime 60 di ${total}` : '';
  list.innerHTML = recent.map((s) => {
    const st = !session.excel ? '' : excelIndex.has(s.serial) ? 'ok' : 'unknown';
    const model = excelIndex.get(s.serial)?.model;
    return `<li>
      <span class="dot ${st}"></span>
      <div class="info">
        <div class="sn">${esc(s.serial)}${s.repeats.length ? `<span class="tag">letto ${s.repeats.length + 1}×</span>` : ''}${s.source === 'manual' ? '<span class="tag info">manuale</span>' : ''}</div>
        <div class="meta">${fmtTime(s.ts)}${model ? ' · ' + esc(model) : ''}${st === 'unknown' ? ' · non in elenco' : ''}</div>
      </div>
      <button class="del" data-sn="${esc(s.serial)}" aria-label="Elimina ${esc(s.serial)}">×</button>
    </li>`;
  }).join('');
}

function setPauseUI() {
  const b = $('#btn-pause');
  b.classList.toggle('on', scan.paused);
  b.innerHTML = scan.paused
    ? '<svg viewBox="0 0 24 24" width="22" height="22"><path fill="currentColor" d="M8 5v14l11-7z"/></svg>'
    : '<svg viewBox="0 0 24 24" width="22" height="22"><path fill="currentColor" d="M7 5h4v14H7zM13 5h4v14h-4z"/></svg>';
  $('#cam-msg').hidden = !scan.paused;
  if (scan.paused) { $('#cam-msg').textContent = 'In pausa'; clearOverlay(); }
}

function initScan() {
  $('#scan-list').addEventListener('click', (e) => {
    const b = e.target.closest('.del');
    if (b && confirm(`Eliminare la scansione ${b.dataset.sn}?`)) removeScan(b.dataset.sn);
  });
  $('#manual-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const v = norm($('#manual-input').value);
    if (!v) return;
    feedback([addScan(v, 'manual', null, 1)]);
    $('#manual-input').value = '';
  });
  $('#btn-pause').addEventListener('click', () => {
    scan.paused = !scan.paused;
    scan.cand.clear(); scan.windowStart = null;
    setPauseUI();
  });
  $('#btn-torch').addEventListener('click', async () => {
    if (!scan.track) return;
    scan.torch = !scan.torch;
    try { await scan.track.applyConstraints({ advanced: [{ torch: scan.torch }] }); } catch { scan.torch = false; }
    $('#btn-torch').classList.toggle('on', scan.torch);
  });
  $('#btn-finish').addEventListener('click', () => { renderReport(); show('report'); });

  document.addEventListener('visibilitychange', () => {
    if ($('#view-scan').hidden) return;
    if (document.hidden) stopCamera(); else startCamera();
  });
}

/* ------------------------------------------------------------------ analisi */

const ISSUE_TYPES = {
  NON_IN_ELENCO:      { sev: 'bad',  title: 'Seriali scansionati non presenti nel file impianto', desc: 'Pannelli trovati in campo ma assenti dall\'elenco: possibile sostituzione non registrata, pannello di un altro impianto o lettura errata.' },
  NON_SCANSIONATO:    { sev: 'bad',  title: 'Seriali del file impianto non trovati in campo', desc: 'Presenti nell\'elenco ma mai scansionati: pannello mancante, sostituito, etichetta illeggibile o non ancora controllato.' },
  SCANSIONE_RIPETUTA: { sev: 'warn', title: 'Seriali letti più volte', desc: 'Lo stesso seriale è stato letto in momenti diversi: può essere una doppia passata o due pannelli con lo stesso seriale.' },
  DUPLICATO_IN_FILE:  { sev: 'warn', title: 'Seriali ripetuti nel file impianto', desc: 'Lo stesso seriale compare su più righe del file Excel.' },
  DATI_FILE_INCOMPLETI: { sev: 'warn', title: 'Righe incomplete nel file impianto', desc: 'Righe del file Excel senza seriale o senza modello.' },
  FORMATO_ANOMALO:    { sev: 'warn', title: 'Seriali con formato anomalo', desc: 'Lunghezza diversa da quella della maggior parte dei seriali: probabile lettura parziale o etichetta di altro tipo. Verificare a vista.' },
  INSERITO_A_MANO:    { sev: 'warn', title: 'Seriali inseriti a mano', desc: 'Non letti da barcode: verificare che la trascrizione sia corretta.' },
};

function dominantLength(serials) {
  const counts = new Map();
  for (const s of serials) counts.set(s.length, (counts.get(s.length) || 0) + 1);
  let best = null, n = 0;
  for (const [len, c] of counts) if (c > n) { best = len; n = c; }
  return serials.length >= 5 && n / serials.length >= 0.6 ? best : null;
}
function hamming(a, b) {
  let d = 0;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i] && ++d > 2) return d;
  return d;
}

function analyze() {
  const ex = session.excel;
  const issues = [];
  const add = (type, serial, model, detail) => issues.push({ type, serial, model: model || '', detail: detail || '' });

  if (ex) {
    const rowsBySerial = new Map();
    for (const r of ex.rows) {
      if (!r.serial) { add('DATI_FILE_INCOMPLETI', '', r.model, `Riga ${r.row}: seriale mancante`); continue; }
      if (!r.model) add('DATI_FILE_INCOMPLETI', r.serial, '', `Riga ${r.row}: modello mancante`);
      if (!rowsBySerial.has(r.serial)) rowsBySerial.set(r.serial, []);
      rowsBySerial.get(r.serial).push(r);
    }
    for (const [sn, rows] of rowsBySerial) {
      if (rows.length < 2) continue;
      const models = [...new Set(rows.map((r) => r.model).filter(Boolean))];
      add('DUPLICATO_IN_FILE', sn, models.join(' / '),
        `Righe ${rows.map((r) => r.row).join(', ')}` + (models.length > 1 ? ' — modelli diversi!' : ''));
    }

    const notScanned = [...excelIndex.values()].filter((r) => !scanIndex.has(r.serial));
    for (const s of session.scans) {
      if (excelIndex.has(s.serial)) continue;
      const similar = notScanned.filter((r) => r.serial.length === s.serial.length && hamming(r.serial, s.serial) <= 1).map((r) => r.serial);
      add('NON_IN_ELENCO', s.serial, '',
        `Scansionato ${fmtDateTime(s.ts)}` + (similar.length ? ` · simile a: ${similar.slice(0, 3).join(', ')} (possibile lettura errata)` : ''));
    }
    for (const r of notScanned) add('NON_SCANSIONATO', r.serial, r.model, `Riga ${r.row} del file`);
  }

  for (const s of session.scans) {
    if (s.repeats.length) {
      add('SCANSIONE_RIPETUTA', s.serial, excelIndex.get(s.serial)?.model,
        `Letto ${s.repeats.length + 1} volte: ${[s.ts, ...s.repeats].map(fmtTime).join(', ')}`);
    }
    if (s.source === 'manual') add('INSERITO_A_MANO', s.serial, excelIndex.get(s.serial)?.model, `Inserito ${fmtDateTime(s.ts)}`);
  }

  const refLen = dominantLength(ex ? [...excelIndex.keys()] : session.scans.map((s) => s.serial));
  if (refLen) {
    for (const s of session.scans) {
      if (s.serial.length !== refLen && !excelIndex.has(s.serial)) {
        add('FORMATO_ANOMALO', s.serial, '', `${s.serial.length} caratteri invece di ${refLen}`);
      }
    }
  }
  return issues;
}

/* ------------------------------------------------------------------ report */

function renderReport() {
  const issues = analyze();
  const ex = session.excel;
  const total = session.scans.length;
  const ok = ex ? session.scans.filter((s) => excelIndex.has(s.serial)).length : null;
  const byType = groupBy(issues, (i) => i.type);
  const n = (t) => byType.get(t)?.length || 0;
  const critical = n('NON_IN_ELENCO') + n('NON_SCANSIONATO');

  $('#rep-plant').textContent = session.plant;
  $('#rep-date').textContent = fmtDateTime(session.startedAt) + (ex ? ' · ' + ex.fileName : ' · senza file impianto');
  $('#rep-load-excel').hidden = !!ex;

  const cards = [
    `<div><b>${total}</b><span>pannelli scansionati</span></div>`,
  ];
  if (ex) {
    cards.push(
      `<div><b>${excelIndex.size}</b><span>attesi da file</span></div>`,
      `<div class="ok"><b>${ok}</b><span>corrispondenti</span></div>`,
      `<div class="bad"><b>${n('NON_IN_ELENCO')}</b><span>non in elenco</span></div>`,
      `<div class="bad"><b>${n('NON_SCANSIONATO')}</b><span>non trovati</span></div>`,
    );
  }
  cards.push(`<div class="warn"><b>${n('SCANSIONE_RIPETUTA')}</b><span>letti più volte</span></div>`);
  const verdict = !ex
    ? `<div class="verdict">Nessun file impianto: ${issues.length ? issues.length + ' segnalazioni' : 'nessuna segnalazione'}.</div>`
    : critical
      ? `<div class="verdict bad">${critical} incongruenze tra campo e file impianto.</div>`
      : `<div class="verdict ok">Tutti i seriali corrispondono al file impianto.</div>`;
  $('#rep-summary').innerHTML = verdict + cards.join('');

  // Riepilogo per modello
  if (ex) {
    const models = new Map();
    for (const r of excelIndex.values()) {
      const k = r.model || '(senza modello)';
      const m = models.get(k) || { exp: 0, found: 0 };
      m.exp++; if (scanIndex.has(r.serial)) m.found++;
      models.set(k, m);
    }
    $('#rep-models').innerHTML = `<details class="issue"><summary>Riepilogo per modello<span class="count">${models.size}</span></summary>
      <div class="tbl-wrap"><table class="tbl"><tr><th>Modello</th><th>Attesi</th><th>Trovati</th><th>Mancanti</th></tr>
      ${[...models].sort().map(([k, m]) => `<tr><td>${esc(k)}</td><td class="num">${m.exp}</td><td class="num">${m.found}</td><td class="num">${m.exp - m.found}</td></tr>`).join('')}
      </table></div></details>`;
  } else {
    $('#rep-models').innerHTML = '';
  }

  const parts = [];
  for (const [type, meta] of Object.entries(ISSUE_TYPES)) {
    const list = byType.get(type);
    if (!list) continue;
    parts.push(`<details class="issue sev-${meta.sev}"${meta.sev === 'bad' ? ' open' : ''}>
      <summary>${esc(meta.title)}<span class="count">${list.length}</span></summary>
      <p class="desc">${esc(meta.desc)}</p>
      <div class="tbl-wrap"><table class="tbl"><tr><th>Seriale</th><th>Modello</th><th>Dettaglio</th></tr>
      ${list.slice(0, 500).map((i) => `<tr><td class="mono">${esc(i.serial) || '—'}</td><td>${esc(i.model)}</td><td>${esc(i.detail)}</td></tr>`).join('')}
      ${list.length > 500 ? `<tr><td colspan="3" class="muted">… altre ${list.length - 500} nel file CSV</td></tr>` : ''}
      </table></div></details>`);
  }
  $('#rep-issues').innerHTML = parts.length
    ? '<h2 style="margin:18px 0 10px">Incongruenze</h2>' + parts.join('')
    : '<div class="card"><p>Nessuna incongruenza rilevata.</p></div>';

  $('#share-row').hidden = !(navigator.canShare && navigator.canShare({ files: [new File([''], 't.csv', { type: 'text/csv' })] }));
  return issues;
}

/* ------------------------------------------------------------------ CSV */

function csv(rows) {
  const cell = (v) => {
    const s = String(v ?? '');
    return /[";\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  // ; come separatore e BOM UTF-8: si apre correttamente in Excel italiano
  return '﻿' + rows.map((r) => r.map(cell).join(';')).join('\r\n') + '\r\n';
}
function baseName() {
  const d = new Date(session.startedAt);
  const ymd = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  return `${session.plant.replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '') || 'impianto'}_${ymd}`;
}
function scansFile() {
  const ex = session.excel;
  const rows = [['N', 'Seriale', 'Modello (da file impianto)', 'Esito', 'Data', 'Ora', 'Origine', 'Lettura n.', 'Barcode nella lettura', 'Letture ripetute', 'Impianto']];
  session.scans.forEach((s, i) => {
    const d = new Date(s.ts);
    rows.push([
      i + 1, s.serial, excelIndex.get(s.serial)?.model || '',
      !ex ? '' : excelIndex.has(s.serial) ? 'OK' : 'NON IN ELENCO',
      fmtDate(d), fmtTime(d),
      s.source === 'manual' ? 'Manuale' : 'Barcode',
      s.group ?? '', s.source === 'manual' ? '' : s.groupSize,
      s.repeats.length, session.plant,
    ]);
  });
  return new File([csv(rows)], `${baseName()}_scansioni.csv`, { type: 'text/csv' });
}
function issuesFile() {
  const rows = [['Tipo', 'Descrizione', 'Seriale', 'Modello', 'Dettaglio', 'Impianto']];
  for (const i of analyze()) rows.push([i.type, ISSUE_TYPES[i.type].title, i.serial, i.model, i.detail, session.plant]);
  return new File([csv(rows)], `${baseName()}_incongruenze.csv`, { type: 'text/csv' });
}
function download(file) {
  const url = URL.createObjectURL(file);
  const a = Object.assign(document.createElement('a'), { href: url, download: file.name });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function initReport() {
  $('#btn-back').addEventListener('click', () => enterScan());
  $('#dl-scans').addEventListener('click', () => download(scansFile()));
  $('#dl-issues').addEventListener('click', () => download(issuesFile()));
  $('#share-files').addEventListener('click', async () => {
    const files = [scansFile(), issuesFile()];
    try { await navigator.share({ files, title: `solarIDcheck – ${session.plant}` }); }
    catch (e) { if (e.name !== 'AbortError') alert('Condivisione non riuscita: ' + e.message); }
  });
  $('#excel-late').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      session.excel = await parseExcel(f);
      rebuildIndexes(); save(); renderReport();
    } catch (err) { alert('File non leggibile: ' + err.message); }
  });
  $('#btn-new').addEventListener('click', () => {
    if (!confirm('Chiudere la sessione? Le scansioni verranno cancellate da questo dispositivo. Hai scaricato i file CSV?')) return;
    clearStore();
    location.reload();
  });
}

/* ------------------------------------------------------------------ utilità */

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
const pad = (n) => String(n).padStart(2, '0');
const fmtDate = (d) => { d = new Date(d); return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`; };
const fmtTime = (d) => { d = new Date(d); return `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`; };
const fmtDateTime = (d) => `${fmtDate(d)} ${fmtTime(d).slice(0, 5)}`;
function groupBy(arr, fn) {
  const m = new Map();
  for (const x of arr) { const k = fn(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); }
  return m;
}

/* ------------------------------------------------------------------ accesso */

// Si salva solo l'impronta SHA-256 della chiave, non la chiave in chiaro.
const ACCESS_HASH = 'd76900b343e8795129e975140868e06a1f022e9d87c5436cc899a9736e3a8446';
const ACCESS_KEY = 'solaridcheck.access';

async function sha256(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function initLock() {
  let granted = false;
  try { granted = localStorage.getItem(ACCESS_KEY) === ACCESS_HASH; } catch {}
  if (granted) { show('setup'); return; }
  show('lock');
  $('#lock-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const ok = (await sha256($('#lock-key').value)) === ACCESS_HASH;
    $('#lock-err').hidden = ok;
    if (!ok) { $('#lock-key').select(); return; }
    try { localStorage.setItem(ACCESS_KEY, ACCESS_HASH); } catch {}
    $('#lock-key').value = '';
    show('setup');
  });
}

/* ------------------------------------------------------------------ avvio */

initLock();
initSetup();
initScan();
initReport();
if ('serviceWorker' in navigator && location.protocol === 'https:') {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
