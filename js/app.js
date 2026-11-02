/* APEX Dashboard - application bootstrap / orchestration */
(function () {
  const { config, utils, theme, data, model, kpi, charts, filters, tables } = APEX;
  const { $, $$, toast, fmtTime, fmtNum } = utils;

  const state = { records: null, model: null, globalSig: null, loaded: false, lastUpdate: null };

  /* ---------- Loader ---------- */
  const loader = {
    el: $('#loader'),
    show(message, { error = false, actions = false, hint = '' } = {}) {
      this.el.classList.remove('hide');
      this.el.classList.toggle('error', error);
      $('#loaderMessage').textContent = message;
      $('#loaderActions').hidden = !actions;
      $('#loaderHint').hidden = !hint;
      $('#loaderHint').innerHTML = hint;
    },
    hide() { this.el.classList.add('hide'); }
  };

  const SERVE_HINT = 'Tip: run <code>start-dashboard.bat</code> in this folder to open the dashboard with fully automatic refresh, or connect the workbook once using the button above.';

  /* ---------- Status pill ---------- */
  function setStatus(kind, text) {
    const pill = $('#statusPill');
    pill.className = `status-pill ${kind}`;
    $('#statusText').textContent = text;
  }

  function liveStatus() {
    if (!state.lastUpdate) return;
    const m = data.getMode();
    const every = `${Math.round(config.pollInterval / 1000)}s`;
    if (m === 'manual') setStatus('warn', `Manual · loaded ${fmtTime(state.lastUpdate)}`);
    else setStatus('live', `Live · updated ${fmtTime(state.lastUpdate)} · checks every ${every}`);
  }

  /* ---------- Tabs ---------- */
  function activateTab(name) {
    $$('.tab').forEach(t => {
      const on = t.dataset.tab === name;
      t.classList.toggle('active', on);
      t.setAttribute('aria-selected', on);
    });
    $$('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `tab-${name}`));
    try { localStorage.setItem(config.storageKeys.tab, name); } catch { /* ignore */ }
    if (name === 'di') requestAnimationFrame(() => charts.resize());
  }

  /* ---------- Rendering ---------- */
  function renderDetail() {
    if (!state.records) return;
    tables.renderDetail(filters.apply(state.records), state.records.length);
  }

  /* KPIs, charts and matrix – scoped by dashboard-wide filters (Assigned To) */
  function renderDashboard() {
    state.globalSig = filters.signature(config.globalFilterKeys);
    const m = state.model = model.build(filters.apply(state.records, config.globalFilterKeys));
    kpi.render(m);
    charts.render(m);
    tables.renderMatrix(m.matrix);
  }

  function renderAll() {
    filters.setOptions(model.filterValues(state.records, config.filterKeys));
    renderDashboard();
    renderDetail();
  }

  function onFiltersChanged() {
    if (!state.records) return;
    if (filters.signature(config.globalFilterKeys) !== state.globalSig) safe(renderDashboard, 'Dashboard');
    renderDetail();
  }

  function safe(fn, label) {
    try { fn(); } catch (err) {
      console.error(`[APEX] ${label} failed`, err);
      toast(`${label} could not be rendered: ${err.message}`, 'error', 7000);
    }
  }

  /* ---------- Data events ---------- */
  data.on('data', ({ records, at }) => {
    state.lastUpdate = at;
    state.records = records;
    safe(renderAll, 'Dashboard');

    if (!state.loaded) {
      state.loaded = true;
      loader.hide();
      if (!records.length) toast('The worksheet contains no data rows.', 'info', 6000);
    } else {
      toast(`Workbook updated · ${fmtNum(records.length)} records refreshed`, 'success');
      $$('#tab-di .card').forEach(c => { c.classList.remove('flash'); void c.offsetWidth; c.classList.add('flash'); });
    }
    liveStatus();
  });

  data.on('status', () => {
    $('#refreshBtn').classList.remove('spinning');
    liveStatus();
  });

  data.on('checking', ({ force }) => { if (force) $('#refreshBtn').classList.add('spinning'); });

  data.on('error', err => {
    console.error('[APEX] data error', err);
    $('#refreshBtn').classList.remove('spinning');
    const msg = err && err.message ? err.message : String(err);
    if (!state.loaded) {
      const fileBlocked = !data.isHttp || /fetch|network/i.test(msg);
      loader.show(msg, { error: true, actions: true, hint: fileBlocked ? SERVE_HINT : '' });
      setStatus('error', 'Data source unavailable');
    } else {
      setStatus('warn', `Last good data ${fmtTime(state.lastUpdate)} · retrying`);
      toast(`Could not read workbook (will retry): ${msg}`, 'error', 6000);
    }
  });

  data.on('needsFile', ({ reconnect, supportsFS }) => {
    const msg = reconnect
      ? 'Click below to re-authorise read access to the workbook.'
      : supportsFS
        ? `Connect "${config.workbook}" once to enable live dashboards. The file is opened read-only and is never modified.`
        : `Select "${config.workbook}" to load the dashboard. The file is opened read-only.`;
    loader.show(msg, { actions: true, hint: SERVE_HINT });
    $('#loaderConnect span').textContent = reconnect ? 'Reconnect Excel File' : 'Connect Excel File';
    setStatus('warn', 'Waiting for workbook');
  });

  /* ---------- UI wiring ---------- */
  function wire() {
    $('#year').textContent = new Date().getFullYear();
    $('#sourceInfo').textContent = `Source: ${config.workbook} › ${config.sheet} (read-only)`;

    $$('.tab').forEach(t => t.addEventListener('click', () => activateTab(t.dataset.tab)));
    let saved = 'di';
    try { saved = localStorage.getItem(config.storageKeys.tab) || 'di'; } catch { /* ignore */ }
    activateTab($(`.tab[data-tab="${saved}"]`) ? saved : 'di');

    $('#refreshBtn').addEventListener('click', () => (data.getMode() && data.getMode() !== 'manual' ? data.refresh() : data.connect()));
    $('#connectBtn').addEventListener('click', () => data.connect({ pickNew: true }));
    $('#loaderConnect').addEventListener('click', () => {
      loader.show('Reading workbook…');
      data.connect().then(r => { if (r === 'aborted' && !state.loaded) data.start(); });
    });
    $('#loaderRetry').addEventListener('click', () => { loader.show('Loading workbook data…'); data.start(); });

    theme.onChange(() => { if (state.model) safe(() => charts.render(state.model), 'Charts'); });

    filters.init(onFiltersChanged);

    charts.onManagerClick(({ category, manager }) => {
      filters.set({ category: [category], manager: [manager] });
      $('#detailCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
      toast(`Filtered detail records: ${category} · ${manager}`, 'info', 3000);
    });
  }

  /* ---------- Boot ---------- */
  function missingLibraries() {
    const libs = { 'jQuery': window.jQuery, 'SheetJS (xlsx)': window.XLSX, 'Chart.js': window.Chart, 'DataTables': window.DataTable };
    return Object.entries(libs).filter(([, v]) => !v).map(([k]) => k);
  }

  theme.init();
  wire();

  const missing = missingLibraries();
  if (missing.length) {
    loader.show(`Required libraries failed to load: ${missing.join(', ')}. Check your network / CDN access and reload.`, { error: true });
    $('#loaderActions').hidden = true;
    setStatus('error', 'Libraries unavailable');
    return;
  }

  loader.show('Loading workbook data…');
  data.start().catch(err => {
    loader.show(err.message || String(err), { error: true, actions: true, hint: SERVE_HINT });
  });
})();
