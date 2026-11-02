/* APEX Dashboard - PREPAY tab orchestration (lazy-loaded when the PREPAY tab is first opened) */
(function () {
  const { utils, prepayConfig: cfg, prepayData: D, prepayCharts: PC, prepayReports: PR, filters, theme } = APEX;
  const { $, $$, esc, fmtNum, fmtDate, fmtTime, toast, animateCount } = utils;
  if (!window.XLSX || !window.Chart || !window.DataTable) return;

  const state = { all: null, sheets: [], model: null, started: false, loaded: false, last: null, file: null };
  const ms = {};
  const source = APEX.createWorkbookSource({ files: cfg.files, idbKey: cfg.idbKey, pollInterval: cfg.pollInterval, parse: D.parse });

  /* ---------- state card ---------- */
  function showState(title, msg, actions = false) {
    $('#ppState').hidden = false;
    $('#ppContent').hidden = true;
    $('#ppStateTitle').textContent = title;
    $('#ppStateMsg').textContent = msg;
    $('#ppStateActions').hidden = !actions;
  }

  function setStatus(kind, text) {
    $('#ppStatusPill').className = `status-pill ${kind}`;
    $('#ppStatusText').textContent = text;
  }

  const liveStatus = () => state.last && setStatus(source.getMode() === 'manual' ? 'warn' : 'live',
    source.getMode() === 'manual' ? `Manual · loaded ${fmtTime(state.last)}` : `Live · updated ${fmtTime(state.last)}`);

  /* ---------- filters ---------- */
  const selections = () => Object.fromEntries(Object.entries(ms).map(([k, m]) => [k, m.values()]));

  function renderChips() {
    const labels = Object.fromEntries(cfg.filters.map(f => [f.key, f.label]));
    $('#ppFilterChips').innerHTML = Object.entries(selections())
      .flatMap(([k, set]) => [...set].map(v => `<span class="filter-chip">${esc(labels[k])}: ${esc(v)}<button type="button" data-key="${k}" data-val="${esc(v)}" aria-label="Remove filter">&times;</button></span>`))
      .join('');
  }

  /* ---------- KPIs ---------- */
  const pct = (a, b) => (b ? `${((a / b) * 100).toFixed(1)}%` : '0%');

  function renderKpis(k) {
    ['total', 'finalized', 'valid', 'open', 'high'].forEach(key => animateCount($(`[data-ppkpi="${key}"]`), k[key]));
    $('[data-ppkpi="avgDm"]').textContent = k.avgDm === null ? '–' : k.avgDm.toFixed(1);
    const sub = (key, text) => { $(`[data-ppkpi-sub="${key}"]`).textContent = text; };
    sub('total', `${fmtNum(k.unique)} unique DCNs${k.latest ? ` · latest ${k.latest.label}: ${fmtNum(k.latest.total)}` : ''}`);
    sub('finalized', `${pct(k.finalized, k.total)} of records`);
    sub('valid', `${pct(k.valid, k.total)} verified valid`);
    sub('open', `${pct(k.open, k.total)} not finalized`);
    sub('avgDm', k.avgDm === null ? 'No DM Age data' : `days · max ${fmtNum(k.maxDm)} days`);
    sub('high', `Urgency ≥ ${cfg.highUrgency} · ${pct(k.high, k.total)}`);
  }

  /* ---------- render ---------- */
  function render() {
    if (!state.all) return;
    try {
      const recs = D.applyFilters(state.all, selections());
      state.model = D.build(recs);
      renderKpis(state.model.kpi);
      PC.render(state.model);
      PR.render(recs);
      renderChips();
    } catch (err) {
      console.error('[APEX PREPAY] render failed', err);
      toast(`PREPAY dashboard could not be rendered: ${err.message}`, 'error', 7000);
    }
  }

  /* ---------- data events ---------- */
  source.on('data', ({ records, sheets, skipped, jira, jiraSheet, at, file }) => {
    state.all = records;
    state.sheets = sheets;
    state.last = at;
    state.file = file;

    const opts = D.filterOptions(records);
    Object.entries(ms).forEach(([k, m]) => m.setOptions(opts[k] || []));

    $('#ppState').hidden = true;
    $('#ppContent').hidden = false;
    const range = sheets.length ? `${fmtDate(sheets[0].date)} – ${fmtDate(sheets[sheets.length - 1].date)}` : '';
    $('#ppSourceInfo').textContent = `${file} · ${sheets.length} October snapshot sheets (${range}) · ${fmtNum(records.length)} records${skipped.length ? ` · skipped: ${skipped.join(', ')}` : ''}`;

    render();
    APEX.prepayJira.load(jira, jiraSheet);
    requestAnimationFrame(() => PC.resize());
    if (state.loaded) toast(`PREPAY workbook updated · ${fmtNum(records.length)} records refreshed`, 'success');
    state.loaded = true;
    liveStatus();
  });

  source.on('status', () => { $('#ppRefresh').classList.remove('spinning'); liveStatus(); });

  source.on('error', err => {
    console.error('[APEX PREPAY] data error', err);
    $('#ppRefresh').classList.remove('spinning');
    const msg = err && err.message ? err.message : String(err);
    if (!state.loaded) showState('PREPAY Data Unavailable', msg, true);
    else {
      setStatus('warn', `Last good data ${fmtTime(state.last)} · retrying`);
      toast(`Could not read PREPAY workbook (will retry): ${msg}`, 'error', 6000);
    }
  });

  source.on('needsFile', ({ reconnect }) => showState(
    'Connect PREPAY Workbook',
    reconnect ? 'Click below to re-authorise read access to the PREPAY workbook.' : `Connect "${cfg.files[0]}" once to enable the PREPAY dashboards. The file is opened read-only and is never modified.`,
    true
  ));

  /* ---------- wiring ---------- */
  function ensureStarted() {
    if (state.started) return;
    state.started = true;
    source.start().catch(err => showState('PREPAY Data Unavailable', err.message || String(err), true));
  }

  $$('[data-pfilter]').forEach(el => { ms[el.dataset.pfilter] = filters.createMultiSelect(el, render); });

  $('#ppResetFilters').addEventListener('click', () => { Object.values(ms).forEach(m => m.clear()); render(); });
  $('#ppFilterChips').addEventListener('click', e => {
    const b = e.target.closest('button[data-key]');
    if (!b) return;
    ms[b.dataset.key].remove(b.dataset.val);
    render();
  });

  $('#ppRefresh').addEventListener('click', () => {
    if (source.getMode() && source.getMode() !== 'manual') { $('#ppRefresh').classList.add('spinning'); source.refresh(); }
    else source.connect();
  });
  $('#ppConnect').addEventListener('click', () => {
    showState('Loading PREPAY Data', 'Reading the PREPAY workbook…');
    source.connect().then(r => { if (r === 'aborted' && !state.loaded) { state.started = false; ensureStarted(); } });
  });
  $('#ppRetry').addEventListener('click', () => {
    showState('Loading PREPAY Data', 'Reading the PREPAY workbook…');
    state.started = false;
    ensureStarted();
  });

  PC.onDrill((key, value) => {
    if (!ms[key]) return;
    ms[key].set([value]);
    render();
    const label = (cfg.filters.find(f => f.key === key) || {}).label || key;
    toast(`PREPAY filtered: ${label} = ${value}`, 'info', 3000);
  });

  theme.onChange(() => { if (state.model) PC.render(state.model); });

  PR.init();
  APEX.prepayJira.init();

  $$('.tab').forEach(t => t.addEventListener('click', () => {
    if (t.dataset.tab !== 'prepay') return;
    ensureStarted();
    requestAnimationFrame(() => PC.resize());
  }));
  if ($('#tab-prepay').classList.contains('active')) ensureStarted();
})();
