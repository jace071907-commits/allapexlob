/* APEX Dashboard - multi-select dropdown filters for the Detail Records table */
window.APEX = window.APEX || {};

APEX.filters = (() => {
  const { $, $$, esc, fmtNum } = APEX.utils;
  const { blankLabel, filterKeys, globalFilterKeys } = APEX.config;
  const instances = {};
  let changeHandler = () => {};

  const CARET = '<svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg>';

  function createMultiSelect(root, notify = () => changeHandler()) {
    const key = root.dataset.filter || root.dataset.pfilter || root.dataset.jfilter || root.dataset.nfilter || root.dataset.ifilter;
    const selected = new Set();
    let options = [];

    root.innerHTML = `
      <button class="ms-toggle" type="button" aria-haspopup="listbox" aria-expanded="false">
        <span class="ms-label">${esc(root.dataset.label)}</span>
        <span class="ms-value">All</span>${CARET}
      </button>
      <div class="ms-panel" role="listbox" aria-multiselectable="true">
        <input class="ms-search" type="search" placeholder="Search ${esc(root.dataset.label)}…" />
        <div class="ms-actions"><button type="button" data-act="all">Select all</button><button type="button" data-act="none">Clear Selection</button></div>
        <ul class="ms-list"></ul>
      </div>`;

    const toggle = $('.ms-toggle', root);
    const search = $('.ms-search', root);
    const list = $('.ms-list', root);

    const summary = () => {
      const n = selected.size;
      const txt = !n || n === options.length ? 'All' : n === 1 ? [...selected][0] : `${n} selected`;
      $('.ms-value', root).textContent = txt;
      root.classList.toggle('active', n > 0 && n < options.length);
    };

    const draw = () => {
      const q = search.value.trim().toLowerCase();
      const shown = options.filter(o => o.name.toLowerCase().includes(q));
      list.innerHTML = shown.length
        ? shown.map(o => `<li><label><input type="checkbox" value="${esc(o.name)}" ${selected.has(o.name) ? 'checked' : ''}/><span class="ms-text">${esc(o.name)}</span><span class="ms-count">${fmtNum(o.count)}</span></label></li>`).join('')
        : '<li class="ms-empty">No matching values</li>';
    };

    const emit = () => { summary(); notify(); };

    toggle.addEventListener('click', e => {
      e.stopPropagation();
      const open = !root.classList.contains('open');
      $$('.multi-select.open').forEach(m => m.classList.remove('open'));
      root.classList.toggle('open', open);
      toggle.setAttribute('aria-expanded', open);
      if (open) { search.value = ''; draw(); setTimeout(() => search.focus(), 50); }
    });
    root.addEventListener('click', e => e.stopPropagation());
    search.addEventListener('input', draw);
    list.addEventListener('change', e => {
      if (e.target.type !== 'checkbox') return;
      e.target.checked ? selected.add(e.target.value) : selected.delete(e.target.value);
      emit();
    });
    $('.ms-actions', root).addEventListener('click', e => {
      const act = e.target.dataset.act;
      if (!act) return;
      selected.clear();
      if (act === 'all') options.forEach(o => selected.add(o.name));
      draw();
      emit();
    });

    return {
      key,
      label: root.dataset.label,
      setOptions(opts) {
        options = opts;
        const valid = new Set(opts.map(o => o.name));
        [...selected].forEach(v => { if (!valid.has(v)) selected.delete(v); });
        draw();
        summary();
      },
      set(values) { selected.clear(); values.forEach(v => selected.add(v)); draw(); summary(); },
      remove(v) { selected.delete(v); draw(); summary(); },
      clear() { selected.clear(); draw(); summary(); },
      values: () => (selected.size === options.length ? new Set() : new Set(selected))
    };
  }

  function renderChips() {
    const chips = [];
    Object.values(instances).forEach(inst => {
      inst.values().forEach(v => chips.push(`<span class="filter-chip">${esc(inst.label)}: ${esc(v)}<button type="button" data-key="${inst.key}" data-val="${esc(v)}" aria-label="Remove filter">&times;</button></span>`));
    });
    $('#filterChips').innerHTML = chips.join('');
  }

  function init(onChange) {
    changeHandler = () => { renderChips(); onChange(); };
    $$('.multi-select[data-filter]').forEach(el => { instances[el.dataset.filter] = createMultiSelect(el); });
    document.addEventListener('click', () => $$('.multi-select.open').forEach(m => m.classList.remove('open')));
    document.addEventListener('keydown', e => { if (e.key === 'Escape') $$('.multi-select.open').forEach(m => m.classList.remove('open')); });
    $('#resetFilters')?.addEventListener('click', () => { Object.values(instances).forEach(i => i.clear()); changeHandler(); });
    $('#filterChips')?.addEventListener('click', e => {
      const b = e.target.closest('button[data-key]');
      if (!b) return;
      instances[b.dataset.key].remove(b.dataset.val);
      changeHandler();
    });
  }

  function setOptions(values) {
    filterKeys.forEach(k => instances[k] && instances[k].setOptions(values[k] || []));
    renderChips();
  }

  /* Replaces table-level selections; dashboard-wide filters (e.g. Assigned To) are kept */
  function set(selection) {
    Object.values(instances).forEach(i => { if (!globalFilterKeys.includes(i.key)) i.clear(); });
    Object.entries(selection).forEach(([k, vals]) => instances[k] && instances[k].set(vals));
    changeHandler();
  }

  function apply(records, keys = filterKeys) {
    const active = keys.map(k => [k, instances[k] ? instances[k].values() : new Set()]).filter(([, s]) => s.size);
    if (!active.length) return records;
    return records.filter(r => active.every(([k, s]) => s.has(r[k] || blankLabel)));
  }

  const signature = keys => keys.map(k => (instances[k] ? [...instances[k].values()].sort().join('\u0001') : '')).join('\u0002');

  return { init, setOptions, set, apply, signature, createMultiSelect };
})();
