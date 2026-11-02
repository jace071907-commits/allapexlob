/* APEX Dashboard - workbook reader (READ-ONLY) with automatic change detection.
 *
 * The workbook is never written to. Two read strategies are supported:
 *   1. HTTP   – page served from a web server (e.g. start-dashboard.bat). The file is polled
 *               with HEAD / GET requests and re-parsed when its content changes.
 *   2. Handle – page opened directly from disk (file://). The user grants read access once via
 *               the File System Access API; the handle is remembered (IndexedDB) and the file
 *               is polled for lastModified / size changes.
 *   A plain <input type="file"> fallback is used in browsers without the File System Access API.
 */
window.APEX = window.APEX || {};

APEX.data = (() => {
  const cfg = APEX.config;
  const { str, toDate, hash } = APEX.utils;
  const listeners = {};
  const isHttp = /^https?:$/.test(location.protocol);
  const supportsFS = typeof window.showOpenFilePicker === 'function';

  let mode = null;          // 'http' | 'handle' | 'manual'
  let handle = null;
  let signature = null;     // { head, hash }
  let timer = null;
  let busy = false;

  const on = (evt, fn) => ((listeners[evt] = listeners[evt] || []).push(fn), api);
  const emit = (evt, payload) => (listeners[evt] || []).forEach(fn => fn(payload));

  /* ---------- tiny IndexedDB key/value store for the file handle ---------- */
  const idb = (() => {
    const open = () => new Promise((res, rej) => {
      const r = indexedDB.open('apex-dashboard', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    const tx = async (type, fn) => {
      const db = await open();
      return new Promise((res, rej) => {
        const t = db.transaction('kv', type);
        const req = fn(t.objectStore('kv'));
        t.oncomplete = () => res(req && req.result);
        t.onerror = () => rej(t.error);
      });
    };
    return {
      get: k => tx('readonly', s => s.get(k)).catch(() => null),
      set: (k, v) => tx('readwrite', s => s.put(v, k)).catch(() => null)
    };
  })();

  /* ---------- parsing ---------- */
  function resolveColumns(headers) {
    const lookup = new Map(headers.map(h => [str(h).toLowerCase(), h]));
    const map = {};
    Object.entries(cfg.columns).forEach(([key, aliases]) => {
      const found = aliases.find(a => lookup.has(a.toLowerCase()));
      if (found) map[key] = lookup.get(found.toLowerCase());
    });
    const missing = cfg.requiredColumns.filter(k => !map[k]);
    if (missing.length) {
      throw new Error(`Worksheet "${cfg.sheet}" is missing required column(s): ${missing.map(k => cfg.columns[k][0]).join(', ')}`);
    }
    return map;
  }

  function parse(buffer) {
    if (typeof XLSX === 'undefined') throw new Error('SheetJS library failed to load. Check network access to the CDN.');
    const wb = XLSX.read(buffer, { type: 'array', sheets: cfg.sheet, cellDates: false });
    const ws = wb.Sheets[cfg.sheet];
    if (!ws) throw new Error(`Worksheet "${cfg.sheet}" was not found in the workbook.`);

    const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
    if (!rows.length) return [];

    const map = resolveColumns(Object.keys(rows[0]));
    const keys = Object.keys(map);

    return rows
      .map((row, i) => {
        const rec = { _row: i + 2 };
        keys.forEach(k => { rec[k] = row[map[k]]; });
        keys.forEach(k => {
          if (k === 'date') rec.date = toDate(rec.date);
          else if (k === 'pendedDays') rec.pendedDays = rec.pendedDays === '' ? '' : Number(rec.pendedDays);
          else rec[k] = str(rec[k]);
        });
        return rec;
      })
      .filter(r => r.category || r.manager || r.inquiry || r.subStatus || r.status);
  }

  /* ---------- sources ---------- */
  const headSig = res => {
    const lm = res.headers.get('last-modified');
    return lm ? `${lm}|${res.headers.get('content-length')}` : null;
  };

  async function readHttp(force) {
    const url = encodeURI(cfg.workbook);
    if (!force && signature && signature.head) {
      try {
        const h = await fetch(url, { method: 'HEAD', cache: 'no-store' });
        if (h.ok && headSig(h) === signature.head) return null;
      } catch { /* fall through to full GET */ }
    }
    const res = await fetch(`${url}?_=${Date.now()}`, { cache: 'no-store' });
    if (!res.ok) throw new Error(`Unable to load "${cfg.workbook}" (HTTP ${res.status}).`);
    const buf = await res.arrayBuffer();
    const sig = { head: headSig(res), hash: hash(new Uint8Array(buf)) };
    const unchanged = signature && signature.hash === sig.hash;
    signature = sig;
    return unchanged && !force ? null : buf;
  }

  async function readHandle(force) {
    const file = await handle.getFile();
    const head = `${file.lastModified}|${file.size}`;
    if (!force && signature && signature.head === head) return null;
    const buf = await file.arrayBuffer();
    const sig = { head, hash: hash(new Uint8Array(buf)) };
    const unchanged = signature && signature.hash === sig.hash;
    signature = sig;
    return unchanged && !force ? null : buf;
  }

  /* ---------- polling ---------- */
  async function check(force = false) {
    if (busy || !mode || mode === 'manual') return;
    busy = true;
    emit('checking', { force });
    try {
      const buf = mode === 'http' ? await readHttp(force) : await readHandle(force);
      if (buf) emit('data', { records: parse(buf), at: new Date(), mode });
      emit('status', { ok: true, at: new Date(), mode });
    } catch (err) {
      // a save in progress can produce a partially written file – keep last good data and retry next poll
      if (signature) signature.head = null;
      emit('error', err);
    } finally {
      busy = false;
    }
  }

  function schedule() {
    clearInterval(timer);
    timer = setInterval(() => check(false), cfg.pollInterval);
  }

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') check(false);
  });

  async function useHandle(h) {
    handle = h;
    mode = 'handle';
    signature = null;
    await check(true);
    schedule();
  }

  async function start() {
    if (isHttp) {
      mode = 'http';
      await check(true);
      schedule();
      return;
    }
    if (supportsFS) {
      const stored = await idb.get('handle');
      if (stored && (await stored.queryPermission({ mode: 'read' })) === 'granted') return useHandle(stored);
      emit('needsFile', { reconnect: !!stored, supportsFS });
      return;
    }
    emit('needsFile', { reconnect: false, supportsFS });
  }

  /* Must be called from a user gesture (click) */
  async function connect({ pickNew = false } = {}) {
    if (supportsFS) {
      try {
        const stored = pickNew ? null : await idb.get('handle');
        if (stored && (await stored.requestPermission({ mode: 'read' })) === 'granted') return useHandle(stored);
        const [picked] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] } }]
        });
        await idb.set('handle', picked);
        return useHandle(picked);
      } catch (err) {
        if (err && err.name === 'AbortError') return 'aborted';
        emit('error', err);
      }
      return;
    }
    const input = Object.assign(document.createElement('input'), { type: 'file', accept: '.xlsx,.xlsm' });
    input.addEventListener('change', async () => {
      const file = input.files && input.files[0];
      if (!file) return;
      try {
        mode = 'manual';
        emit('data', { records: parse(await file.arrayBuffer()), at: new Date(), mode });
        emit('status', { ok: true, at: new Date(), mode });
      } catch (err) { emit('error', err); }
    });
    input.click();
  }

  const refresh = () => check(true);
  const getMode = () => mode;

  const api = { on, start, connect, refresh, getMode, isHttp, supportsFS };
  return api;
})();
