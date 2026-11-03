/* APEX Dashboard - CASH CONTROL tab orchestration (lazy-loaded when the CASH CONTROL tab is first opened)
 * Pipeline: all records → filters + global search → KPIs, charts and report table
 */
(function () {
  const { utils, cashConfig: cfg, cashData: D, filters, theme } = APEX;
  const { $, $$, esc, fmtNum, fmtTime, toast, animateCount, debounce } = utils;
  if (!window.XLSX || !window.Chart || !window.DataTable) return;

  const state = { all: null, model: null, started: false, loaded: false, last: null, query: '' };
  const ms = {};
  let dt = null;
  const source = APEX.createWorkbookSource({ files: cfg.files, idbKey: cfg.idbKey, pollInterval: cfg.pollInterval, parse: D.parse, accept: ['.xlsb', '.xlsx', '.xlsm'] });

  /* ---------- helpers ---------- */
  function showState(title, msg, actions = false) {
    $('#ccState').hidden = false;
    $('#ccContent').hidden = true;
    $('#ccStateTitle').textContent = title;
    $('#ccStateMsg').textContent = msg;
    $('#ccStateActions').hidden = !actions;
  }
  const setStatus = (kind, text) => { $('#ccStatusPill').className = `status-pill ${kind}`; $('#ccStatusText').textContent = text; };
  const liveStatus = () => state.last && setStatus(source.getMode() === 'manual' ? 'warn' : 'live', `${source.getMode() === 'manual' ? 'Manual · loaded' : 'Live · updated'} ${fmtTime(state.last)}`);
  const selections = () => Object.fromEntries(Object.entries(ms).map(([k, m]) => [k, m.values()]));
  const pct = (a, b) => (b ? `${((a / b) * 100).toFixed(1)}%` : '0%');
  const badge = (id, text) => { const el = $(`#${id}`); if (el) el.textContent = text; };

  /* ---------- KPIs ---------- */
  function renderKpis(k) {
    const set = (key, v, sub) => { animateCount($(`[data-ckpi="${key}"]`), v); $(`[data-ckpi-sub="${key}"]`).textContent = sub; };
    set('total', k.total, k.unique === k.total ? 'All COR numbers unique' : `${fmtNum(k.unique)} unique COR numbers`);
    set('managers', k.managers, k.managers ? `${fmtNum(Math.round(k.total / k.managers))} CORs per manager` : '–');
    set('associates', k.associates, k.associates ? `${fmtNum(Math.round(k.total / k.associates))} CORs per associate` : '–');
    set('domains', k.domains, 'Distinct associate domains');
    set('a0', k.a0, `${pct(k.a0, k.total)} of CORs`);
    set('a1', k.a1, `${pct(k.a1, k.total)} of CORs`);
    set('a2', k.a2, `${pct(k.a2, k.total)} of CORs`);
    set('a3', k.a3, `${pct(k.a3, k.total)} of CORs${k.noAge ? ` · ${fmtNum(k.noAge)} without age` : ''}`);
  }

  /* ---------- charts ---------- */
  function doughnut(id, rows, c, key) {
    const PC = APEX.prepayCharts, H = APEX.charts.helpers;
    PC.setEmpty(id, !rows.length);
    const palette = [c.accent, c.accent2, c.good, '#ffb547', c.bad, '#8d97ab', '#b48cff', '#4dd0e1'];
    const total = rows.reduce((a, r) => a + r.count, 0);
    PC.upsert(id, {
      type: 'doughnut',
      data: {
        labels: rows.map(r => r.name),
        datasets: [{
          data: rows.map(r => r.count),
          backgroundColor: rows.map((_, i) => palette[i % palette.length]),
          borderColor: c.surface, borderWidth: 3, hoverOffset: 12,
          datalabels: {
            display: ctx => total && ctx.dataset.data[ctx.dataIndex] / total >= 0.03,
            color: '#fff', font: { weight: '700', size: 11.5 }, textAlign: 'center',
            textShadowColor: 'rgba(0,0,0,.5)', textShadowBlur: 4,
            formatter: v => `${fmtNum(v)}\n${((v / total) * 100).toFixed(0)}%`
          }
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '58%',
        animation: { animateRotate: true, duration: 1200, easing: 'easeOutQuart' },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length) pick(key, rows[els[0].index].name); },
        plugins: {
          legend: { position: 'right', labels: { color: c.text, usePointStyle: true, boxWidth: 9, padding: 14, font: { weight: '600' } } },
          tooltip: {
            ...H.tooltip(c),
            callbacks: { label: item => ` ${item.label}: ${fmtNum(item.raw)} (${total ? ((item.raw / total) * 100).toFixed(1) : 0}%)`, footer: () => 'Click to filter' },
            footerColor: c.muted, footerFont: { size: 11, style: 'italic', weight: '400' }
          }
        }
      }
    });
  }

  const bar = (id, rows, c, key, opts = {}) => APEX.prepayCharts.bar(id, rows, c, { label: 'CORs', ...opts, onPick: v => pick(key, v) });

  function renderCharts(m) {
    if (!APEX.prepayCharts) return;
    const c = theme.colors();
    doughnut('ccMgrChart', m.byManager, c, 'manager');
    bar('ccAssocChart', m.byAssociate, c, 'associate', { alt: true });
    bar('ccDomainChart', m.byDomain, c, 'domain');
    bar('ccAgeChart', m.byAge, c, 'ageCat', { horizontal: false });
    bar('ccTopMgrChart', m.topManagers, c, 'manager', { alt: true });
    bar('ccTopAssocChart', m.topAssociates, c, 'associate');
    badge('ccMgrBadge', `${m.byManager.length} managers`);
    badge('ccAssocBadge', `${m.byAssociate.length} associates`);
    badge('ccDomainBadge', `${m.byDomain.length} domains`);
    badge('ccAgeBadge', `${m.byAge.length} categories`);
  }

  /* ---------- report ---------- */
  const R = {
    text: (d, t) => (t === 'display' ? esc(d) : d ?? ''),
    age: (d, t, row) => (t === 'sort' || t === 'type' ? (row.ageLow ?? 1e9) : t === 'display' ? esc(d) : d)
  };

  function renderTable(rows) {
    badge('ccTableCount', `${fmtNum(rows.length)} COR${rows.length === 1 ? '' : 's'}`);
    if (dt) { dt.clear().rows.add(rows).draw(false); return; }
    const { exportButtons, layout, language } = APEX.tables.helpers;
    dt = new DataTable('#ccTable', {
      data: rows,
      columns: [
        { data: 'cor', title: 'COR Number', render: R.text },
        { data: 'associate', title: 'Associate Name', render: R.text },
        { data: 'domain', title: 'Associate Domain', render: R.text },
        { data: 'manager', title: 'Manager Name', render: R.text },
        { data: 'ageCat', title: 'Age Category', type: 'num', render: R.age }
      ].map(col => ({ ...col, defaultContent: '' })),
      order: [[4, 'desc']],
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100, { label: 'All', value: -1 }],
      deferRender: true,
      autoWidth: false,
      layout,
      language: { ...language, searchPlaceholder: 'Search report…', emptyTable: 'No CORs match the current search and filters' },
      buttons: exportButtons('CASH CONTROL COR Report')
    });
  }

  function renderChips() {
    const labels = Object.fromEntries(cfg.filters.map(f => [f.key, f.label]));
    const chips = Object.entries(selections()).flatMap(([k, s]) => [...s].map(v =>
      `<span class="filter-chip">${esc(labels[k])}: ${esc(v)}<button type="button" data-key="${k}" data-val="${esc(v)}" aria-label="Remove filter">&times;</button></span>`));
    if (state.query) chips.push(`<span class="filter-chip">Search: “${esc(state.query)}”<button type="button" data-search="1" aria-label="Clear search">&times;</button></span>`);
    $('#ccFilterChips').innerHTML = chips.join('');
  }

  function render() {
    if (!state.all) return;
    try {
      const recs = D.applySearch(D.applySelections(state.all, selections()), state.query);
      state.model = D.build(recs);
      renderKpis(state.model.kpi);
      renderCharts(state.model);
      renderTable(recs);
      renderChips();
      badge('ccSearchCount', recs.length === state.all.length ? `${fmtNum(state.all.length)} CORs` : `${fmtNum(recs.length)} of ${fmtNum(state.all.length)} CORs`);
    } catch (err) {
      console.error('[APEX CASH] render failed', err);
      toast(`CASH CONTROL dashboard could not be rendered: ${err.message}`, 'error', 7000);
    }
  }

  function pick(key, value) {
    if (!ms[key]) return;
    ms[key].set([value]);
    render();
    toast(`CASH CONTROL filtered: ${(cfg.filters.find(f => f.key === key) || {}).label || key} = ${value}`, 'info', 3000);
  }

  /* ---------- data events ---------- */
  source.on('data', ({ records, sheet, at, file }) => {
    state.all = records;
    state.last = at;
    cfg.filters.forEach(f => ms[f.key] && ms[f.key].setOptions(D.options(records, f.key)));
    $('#ccState').hidden = true;
    $('#ccContent').hidden = false;
    $('#ccSourceInfo').textContent = `${file} › ${sheet} · ${fmtNum(records.length)} CORs`;
    render();
    if (state.loaded) toast(`CASH workbook updated · ${fmtNum(records.length)} CORs refreshed`, 'success');
    state.loaded = true;
    liveStatus();
  });
  source.on('status', () => { $('#ccRefresh').classList.remove('spinning'); liveStatus(); });
  source.on('error', err => {
    console.error('[APEX CASH] data error', err);
    $('#ccRefresh').classList.remove('spinning');
    const msg = err && err.message ? err.message : String(err);
    if (!state.loaded) showState('CASH CONTROL Data Unavailable', msg, true);
    else { setStatus('warn', `Last good data ${fmtTime(state.last)} · retrying`); toast(`Could not read CASH workbook (will retry): ${msg}`, 'error', 6000); }
  });
  source.on('needsFile', ({ reconnect }) => showState('Connect CASH Workbook',
    reconnect ? 'Click below to re-authorise read access to the CASH workbook.' : `Connect "${cfg.files[0]}" once to enable the CASH CONTROL dashboards. The file is opened read-only and is never modified.`, true));

  /* ---------- wiring ---------- */
  function ensureStarted() {
    if (state.started) return;
    state.started = true;
    source.start().catch(err => showState('CASH CONTROL Data Unavailable', err.message || String(err), true));
  }

  $$('[data-cfilter]').forEach(el => { ms[el.dataset.cfilter] = filters.createMultiSelect(el, render); });

  const search = $('#ccSearch');
  const clearSearch = () => { search.value = ''; state.query = ''; $('#ccSearchClear').hidden = true; };
  search.addEventListener('input', debounce(() => { state.query = search.value.trim(); $('#ccSearchClear').hidden = !search.value; render(); }, 120));
  $('#ccSearchClear').addEventListener('click', () => { clearSearch(); render(); search.focus(); });

  $('#ccResetFilters').addEventListener('click', () => { Object.values(ms).forEach(m => m.clear()); clearSearch(); render(); });
  $('#ccFilterChips').addEventListener('click', e => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.search) clearSearch(); else ms[b.dataset.key].remove(b.dataset.val);
    render();
  });
  $('#ccRefresh').addEventListener('click', () => {
    if (source.getMode() && source.getMode() !== 'manual') { $('#ccRefresh').classList.add('spinning'); source.refresh(); }
    else source.connect();
  });
  $('#ccConnect').addEventListener('click', () => {
    showState('Loading CASH CONTROL Data', 'Reading the CASH workbook…');
    source.connect().then(r => { if (r === 'aborted' && !state.loaded) { state.started = false; ensureStarted(); } });
  });
  $('#ccRetry').addEventListener('click', () => { showState('Loading CASH CONTROL Data', 'Reading the CASH workbook…'); state.started = false; ensureStarted(); });

  theme.onChange(() => { if (state.model) renderCharts(state.model); });

  $$('.tab').forEach(t => t.addEventListener('click', () => { if (t.dataset.tab === 'cashcontrol') ensureStarted(); }));
  if ($('#tab-cashcontrol').classList.contains('active')) ensureStarted();
})();
