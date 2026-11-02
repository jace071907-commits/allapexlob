/* APEX Dashboard - NON COB tab orchestration (lazy-loaded when the NON COB tab is first opened) */
(function () {
  const { utils, noncobConfig: cfg, noncobData: D, noncobCharts: NC, noncobReports: NR, filters, theme } = APEX;
  const { $, $$, esc, fmtNum, fmtTime, toast, animateCount } = utils;
  if (!window.XLSX || !window.Chart || !window.DataTable) return;

  const state = { data: null, fields: null, model: null, started: false, loaded: false, last: null };
  const ms = {};
  const source = APEX.createWorkbookSource({ files: cfg.files, idbKey: cfg.idbKey, pollInterval: cfg.pollInterval, parse: D.parse });

  function showState(title, msg, actions = false) {
    $('#ncState').hidden = false;
    $('#ncContent').hidden = true;
    $('#ncStateTitle').textContent = title;
    $('#ncStateMsg').textContent = msg;
    $('#ncStateActions').hidden = !actions;
  }
  const setStatus = (kind, text) => { $('#ncStatusPill').className = `status-pill ${kind}`; $('#ncStatusText').textContent = text; };
  const liveStatus = () => state.last && setStatus(source.getMode() === 'manual' ? 'warn' : 'live',
    `${source.getMode() === 'manual' ? 'Manual · loaded' : 'Live · updated'} ${fmtTime(state.last)}`);

  const selections = () => Object.fromEntries(Object.entries(ms).map(([k, m]) => [k, m.values()]));
  const pct = (a, b) => (b ? `${((a / b) * 100).toFixed(1)}%` : '0%');
  const kpi = (key, value, sub) => {
    const el = $(`[data-nkpi="${key}"]`);
    if (typeof value === 'number' && Number.isInteger(value)) animateCount(el, value); else el.textContent = value;
    const s = $(`[data-nkpi-sub="${key}"]`);
    if (s) s.textContent = sub;
  };

  function renderKpis(m) {
    const t = m.totals, a = m.ageing, o = m.opid;
    kpi('all', t.all, `Ageing ${fmtNum(t.ageing)} · R Status ${fmtNum(t.rstatus)} · OPID ${fmtNum(t.opid)}`);
    kpi('ageing', t.ageing, a.unbucketed ? `${fmtNum(a.unbucketed)} without an age category` : 'Ageing sheet records');
    cfg.ageBuckets.forEach(([, , label], i) => kpi(`b${i}`, a.buckets[label], `${pct(a.buckets[label], t.ageing)} of ageing`));

    kpi('rac', a.pend.rac, `${pct(a.pend.rac, t.ageing)} of ageing`);
    kpi('restoration', a.pend.restoration, `${pct(a.pend.restoration, t.ageing)} of ageing`);
    kpi('restricted', a.pend.restricted, `${pct(a.pend.restricted, t.ageing)} of ageing`);
    kpi('hostPlan', a.pend.hostPlan, a.homePlan ? `+ ${fmtNum(a.homePlan)} waiting on Home Plan` : `${pct(a.pend.hostPlan, t.ageing)} of ageing`);
    $('#ncRTotal').textContent = `${fmtNum(t.rstatus)} R Status records`;

    kpi('opid', t.opid, `${fmtNum(o.byProcessor.length)} processors`);
    kpi('carry', o.carry, `${pct(o.carry, t.opid)} of OPID`);
    kpi('nonCarry', o.nonCarry, `${pct(o.nonCarry, t.opid)} of OPID`);
    kpi('avgAge', o.avgAge === null ? '–' : o.avgAge.toFixed(1), o.avgAge === null ? 'No age data' : `days · max ${fmtNum(o.maxAge)} · ${fmtNum(o.agedCount)} with age`);
  }

  function renderChips() {
    const labels = Object.fromEntries(cfg.filters.map(f => [f.key, f.label]));
    $('#ncFilterChips').innerHTML = Object.entries(selections())
      .flatMap(([k, set]) => [...set].map(v => `<span class="filter-chip">${esc(labels[k])}: ${esc(v)}<button type="button" data-key="${k}" data-val="${esc(v)}" aria-label="Remove filter">&times;</button></span>`))
      .join('');
  }

  function render() {
    if (!state.data) return;
    try {
      const f = D.applyFilters(state.data, state.fields, selections());
      state.model = D.build(f);
      renderKpis(state.model);
      NC.render(state.model);
      NR.render(f);
      renderChips();
    } catch (err) {
      console.error('[APEX NON COB] render failed', err);
      toast(`NON COB dashboard could not be rendered: ${err.message}`, 'error', 7000);
    }
  }

  source.on('data', ({ data, fields, missing, at, file }) => {
    state.data = data;
    state.fields = fields;
    state.last = at;
    const opts = D.filterOptions(data, fields);
    Object.entries(ms).forEach(([k, m]) => m.setOptions(opts[k] || []));
    $('#ncState').hidden = true;
    $('#ncContent').hidden = false;
    $('#ncSourceInfo').textContent = `${file} · Ageing ${fmtNum(data.ageing.length)} · R Status ${fmtNum(data.rstatus.length)} · OPID ${fmtNum(data.opid.length)} records${missing.length ? ` · missing sheet(s): ${missing.join(', ')}` : ''}`;
    render();
    if (state.loaded) toast(`NON COB workbook updated · ${fmtNum(data.ageing.length + data.rstatus.length + data.opid.length)} records refreshed`, 'success');
    state.loaded = true;
    liveStatus();
  });
  source.on('status', () => { $('#ncRefresh').classList.remove('spinning'); liveStatus(); });
  source.on('error', err => {
    console.error('[APEX NON COB] data error', err);
    $('#ncRefresh').classList.remove('spinning');
    const msg = err && err.message ? err.message : String(err);
    if (!state.loaded) showState('NON COB Data Unavailable', msg, true);
    else { setStatus('warn', `Last good data ${fmtTime(state.last)} · retrying`); toast(`Could not read NON COB workbook (will retry): ${msg}`, 'error', 6000); }
  });
  source.on('needsFile', ({ reconnect }) => showState('Connect NON COB Workbook',
    reconnect ? 'Click below to re-authorise read access to the NON COB workbook.' : `Connect "${cfg.files[0]}" once to enable the NON COB dashboards. The file is opened read-only and is never modified.`, true));

  function ensureStarted() {
    if (state.started) return;
    state.started = true;
    source.start().catch(err => showState('NON COB Data Unavailable', err.message || String(err), true));
  }

  $$('[data-nfilter]').forEach(el => { ms[el.dataset.nfilter] = filters.createMultiSelect(el, render); });
  $('#ncResetFilters').addEventListener('click', () => { Object.values(ms).forEach(m => m.clear()); render(); });
  $('#ncFilterChips').addEventListener('click', e => {
    const b = e.target.closest('button[data-key]');
    if (!b) return;
    ms[b.dataset.key].remove(b.dataset.val);
    render();
  });
  $('#ncRefresh').addEventListener('click', () => {
    if (source.getMode() && source.getMode() !== 'manual') { $('#ncRefresh').classList.add('spinning'); source.refresh(); }
    else source.connect();
  });
  $('#ncConnect').addEventListener('click', () => {
    showState('Loading NON COB Data', 'Reading the NON COB workbook…');
    source.connect().then(r => { if (r === 'aborted' && !state.loaded) { state.started = false; ensureStarted(); } });
  });
  $('#ncRetry').addEventListener('click', () => { showState('Loading NON COB Data', 'Reading the NON COB workbook…'); state.started = false; ensureStarted(); });

  NC.onPick((key, value) => {
    if (!ms[key]) return;
    ms[key].set([value]);
    render();
    toast(`NON COB filtered: ${(cfg.filters.find(f => f.key === key) || {}).label || key} = ${value}`, 'info', 3000);
  });
  theme.onChange(() => { if (state.model) NC.render(state.model); });
  NR.init();

  $$('.tab').forEach(t => t.addEventListener('click', () => { if (t.dataset.tab === 'noncob') ensureStarted(); }));
  if ($('#tab-noncob').classList.contains('active')) ensureStarted();
})();
