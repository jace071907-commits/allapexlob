/* APEX Dashboard - theme (mode + accent) management with persistence */
window.APEX = window.APEX || {};

APEX.theme = (() => {
  const { storageKeys: K, themeIcons } = APEX.config;
  const { $, $$ } = APEX.utils;
  const root = document.documentElement;
  const listeners = [];

  const store = {
    get: k => { try { return localStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } }
  };

  const mode = () => root.dataset.mode || 'dark';
  const accent = () => root.dataset.accent || 'blue';

  function syncUi() {
    const icon = $('#themeIcon');
    if (icon) icon.src = mode() === 'light' ? themeIcons.light : themeIcons.dark;
    $$('.swatch').forEach(s => {
      const on = s.dataset.accentChoice === accent();
      s.classList.toggle('active', on);
      s.setAttribute('aria-checked', on);
    });
  }

  function apply(next) {
    if (next.mode) { root.dataset.mode = next.mode; store.set(K.mode, next.mode); }
    if (next.accent) { root.dataset.accent = next.accent; store.set(K.accent, next.accent); }
    syncUi();
    // allow CSS variables to settle before consumers read them
    requestAnimationFrame(() => listeners.forEach(fn => fn({ mode: mode(), accent: accent() })));
  }

  function colors() {
    const s = getComputedStyle(root);
    const v = n => s.getPropertyValue(n).trim();
    return {
      accent: v('--accent'),
      accentRgb: v('--accent-rgb'),
      accent2: v('--accent-2'),
      accent2Rgb: v('--accent-2-rgb'),
      text: v('--text'),
      textStrong: v('--text-strong'),
      muted: v('--muted'),
      grid: v('--grid'),
      surface: v('--surface-1'),
      good: v('--good'),
      bad: v('--bad'),
      mode: mode()
    };
  }

  function init() {
    apply({ mode: store.get(K.mode) || mode(), accent: store.get(K.accent) || accent() });
    $('#modeToggle')?.addEventListener('click', () => apply({ mode: mode() === 'dark' ? 'light' : 'dark' }));
    $$('.swatch').forEach(s => s.addEventListener('click', () => apply({ accent: s.dataset.accentChoice })));
  }

  const onChange = fn => listeners.push(fn);

  return { init, onChange, colors, mode, accent };
})();
