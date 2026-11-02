/* APEX Dashboard - shared helpers */
window.APEX = window.APEX || {};

APEX.utils = (() => {
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const nf = new Intl.NumberFormat('en-US');

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const str = v => (v === null || v === undefined ? '' : String(v).trim());

  const esc = v => str(v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const fmtNum = n => nf.format(Math.round(n || 0));

  const fmtPct = (p, digits = 1) => (p === null || !isFinite(p) ? 'N/A' : `${p > 0 ? '+' : ''}${p.toFixed(digits)}%`);

  /* Excel serial / string / Date -> local Date (date part only) */
  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && isFinite(v)) {
      const utc = new Date(Math.round((v - 25569) * 86400000));
      return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }
    const d = new Date(str(v));
    return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  const fmtDate = d => (d ? `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}` : '');
  const fmtTime = d => d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  const monthKey = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  const monthLabel = key => {
    const [y, m] = key.split('-').map(Number);
    return `${MONTHS[m - 1]} ${y}`;
  };

  /* FNV-1a hash – used to detect workbook content changes */
  function hash(bytes) {
    let h = 0x811c9dc5;
    for (let i = 0; i < bytes.length; i++) {
      h ^= bytes[i];
      h = Math.imul(h, 0x01000193);
    }
    return `${(h >>> 0).toString(16)}-${bytes.length}`;
  }

  /* Animated count-up for KPI values */
  function animateCount(el, to, duration = 1100) {
    if (!el) return;
    const from = Number(el.dataset.value || 0);
    el.dataset.value = to;
    if (from === to) { el.textContent = fmtNum(to); return; }
    const start = performance.now();
    const step = now => {
      const t = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - t, 4);
      el.textContent = fmtNum(from + (to - from) * eased);
      if (t < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  const ICONS = {
    info: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg>',
    success: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>',
    error: '<svg viewBox="0 0 24 24"><path d="M12 3l9.5 17h-19z"/><path d="M12 10v4M12 17h.01"/></svg>'
  };

  function toast(message, type = 'info', timeout = 4500) {
    const stack = $('#toastStack');
    if (!stack) return;
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.innerHTML = `${ICONS[type] || ICONS.info}<span>${esc(message)}</span>`;
    stack.appendChild(el);
    setTimeout(() => {
      el.classList.add('out');
      el.addEventListener('animationend', () => el.remove(), { once: true });
    }, timeout);
  }

  const debounce = (fn, ms = 200) => {
    let t;
    return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  };

  const slug = v => str(v).toLowerCase().replace(/[^a-z0-9]+/g, '-');

  return { $, $$, str, esc, fmtNum, fmtPct, toDate, fmtDate, fmtTime, monthKey, monthLabel, hash, animateCount, toast, debounce, slug };
})();
