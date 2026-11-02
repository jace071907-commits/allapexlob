/* APEX Dashboard - generic READ-ONLY workbook source with change detection.
 * Used by tab modules (e.g. PREPAY). Mirrors the DI data service:
 *   - HTTP mode: polls the file (first existing name in `files`) and re-parses on change
 *   - Handle mode (file://): File System Access API handle, remembered in IndexedDB
 * The workbook is never written to.
 */
window.APEX = window.APEX || {};

APEX.createWorkbookSource = function ({ files, idbKey, pollInterval = 15000, parse }) {
  const { hash } = APEX.utils;
  const listeners = {};
  const isHttp = /^https?:$/.test(location.protocol);
  const supportsFS = typeof window.showOpenFilePicker === 'function';
  let mode = null;
  let handle = null;
  let signature = null;
  let resolvedFile = null;
  let timer = null;
  let busy = false;

  const api = {};
  const emit = (evt, p) => (listeners[evt] || []).forEach(fn => fn(p));

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

  const headSig = res => {
    const lm = res.headers.get('last-modified');
    return lm ? `${lm}|${res.headers.get('content-length')}` : null;
  };

  async function readHttp(force) {
    if (!force && signature && signature.head && resolvedFile) {
      try {
        const h = await fetch(encodeURI(resolvedFile), { method: 'HEAD', cache: 'no-store' });
        if (h.ok && headSig(h) === signature.head) return null;
      } catch { /* fall through */ }
    }
    const candidates = resolvedFile ? [resolvedFile, ...files.filter(f => f !== resolvedFile)] : files;
    let res = null;
    for (const f of candidates) {
      res = await fetch(`${encodeURI(f)}?_=${Date.now()}`, { cache: 'no-store' });
      if (res.ok) { resolvedFile = f; break; }
    }
    if (!res || !res.ok) throw new Error(`Unable to load ${files.map(f => `"${f}"`).join(' or ')} (HTTP ${res ? res.status : 'n/a'}).`);
    const buf = await res.arrayBuffer();
    const sig = { head: headSig(res), hash: hash(new Uint8Array(buf)) };
    const unchanged = signature && signature.hash === sig.hash;
    signature = sig;
    return unchanged && !force ? null : buf;
  }

  async function readHandle(force) {
    const file = await handle.getFile();
    resolvedFile = file.name;
    const head = `${file.lastModified}|${file.size}`;
    if (!force && signature && signature.head === head) return null;
    const buf = await file.arrayBuffer();
    const sig = { head, hash: hash(new Uint8Array(buf)) };
    const unchanged = signature && signature.hash === sig.hash;
    signature = sig;
    return unchanged && !force ? null : buf;
  }

  async function check(force = false) {
    if (busy || !mode || mode === 'manual') return;
    busy = true;
    try {
      const buf = mode === 'http' ? await readHttp(force) : await readHandle(force);
      if (buf) emit('data', { ...parse(buf), at: new Date(), file: resolvedFile, mode });
      emit('status', { ok: true, at: new Date(), mode });
    } catch (err) {
      if (signature) signature.head = null;
      emit('error', err);
    } finally {
      busy = false;
    }
  }

  function schedule() {
    clearInterval(timer);
    timer = setInterval(() => check(false), pollInterval);
  }

  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') check(false); });

  async function useHandle(h) {
    handle = h; mode = 'handle'; signature = null;
    await check(true);
    schedule();
  }

  api.on = (evt, fn) => ((listeners[evt] = listeners[evt] || []).push(fn), api);

  api.start = async () => {
    if (isHttp) { mode = 'http'; await check(true); schedule(); return; }
    if (supportsFS) {
      const stored = await idb.get(idbKey);
      if (stored && (await stored.queryPermission({ mode: 'read' })) === 'granted') return useHandle(stored);
      emit('needsFile', { reconnect: !!stored, supportsFS });
      return;
    }
    emit('needsFile', { reconnect: false, supportsFS });
  };

  api.connect = async ({ pickNew = false } = {}) => {
    if (supportsFS) {
      try {
        const stored = pickNew ? null : await idb.get(idbKey);
        if (stored && (await stored.requestPermission({ mode: 'read' })) === 'granted') return useHandle(stored);
        const [picked] = await window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'] } }]
        });
        await idb.set(idbKey, picked);
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
        resolvedFile = file.name;
        emit('data', { ...parse(await file.arrayBuffer()), at: new Date(), file: file.name, mode });
      } catch (err) { emit('error', err); }
    });
    input.click();
  };

  api.refresh = () => check(true);
  api.getMode = () => mode;
  api.isHttp = isHttp;
  return api;
};
