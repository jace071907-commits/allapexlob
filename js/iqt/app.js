/* APEX Dashboard - IQT tab orchestration (lazy-loaded when the IQT tab is first opened)
 * Pipeline:  all records → global filters + global search → KPIs / charts
 *                                                      → report filters + report search + date range → report table
 */
(function () {
  const { utils, iqtConfig: cfg, iqtData: D, iqtCharts: IC, filters, theme } = APEX;
  const { $, $$, esc, fmtNum, fmtDate, fmtTime, toast, animateCount, debounce } = utils;
  if (!window.XLSX || !window.Chart || !window.DataTable) return;

  const state = { all: null, model: null, started: false, loaded: false, last: null, gQuery: '', tQuery: '', from: null, to: null };
  const gms = {};
  const tms = {};
  let dt = null;
  const source = APEX.createWorkbookSource({ files: cfg.files, idbKey: cfg.idbKey, pollInterval: cfg.pollInterval, parse: D.parse });

  /* ---------- helpers ---------- */
  function showState(title, msg, actions = false) {
    $('#iqState').hidden = false;
    $('#iqContent').hidden = true;
    $('#iqStateTitle').textContent = title;
    $('#iqStateMsg').textContent = msg;
    $('#iqStateActions').hidden = !actions;
  }
  const setStatus = (kind, text) => { $('#iqStatusPill').className = `status-pill ${kind}`; $('#iqStatusText').textContent = text; };
  const liveStatus = () => state.last && setStatus(source.getMode() === 'manual' ? 'warn' : 'live', `${source.getMode() === 'manual' ? 'Manual · loaded' : 'Live · updated'} ${fmtTime(state.last)}`);
  const selections = set => Object.fromEntries(Object.entries(set).map(([k, m]) => [k, m.values()]));
  const pct = (a, b) => (b ? `${((a / b) * 100).toFixed(1)}%` : '0%');
  const parseInputDate = v => { if (!v) return null; const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d); };

  function chips(containerId, set, defs, extra = []) {
    const labels = Object.fromEntries(defs.map(f => [f.key, f.label]));
    $(containerId).innerHTML = Object.entries(selections(set))
      .flatMap(([k, s]) => [...s].map(v => `<span class="filter-chip">${esc(labels[k])}: ${esc(v)}<button type="button" data-key="${k}" data-val="${esc(v)}" aria-label="Remove filter">&times;</button></span>`))
      .concat(extra).join('');
  }

  /* ---------- KPIs ---------- */
  function renderKpis(k) {
    const set = (key, v, sub) => {
      const el = $(`[data-ikpi="${key}"]`);
      if (typeof v === 'number' && Number.isInteger(v)) animateCount(el, v); else el.textContent = v;
      $(`[data-ikpi-sub="${key}"]`).textContent = sub;
    };
    set('total', k.total, `${fmtNum(k.unique)} unique inquiries`);
    set('investigation', k.investigation, `${pct(k.investigation, k.total)} of records`);
    set('review', k.review, `${pct(k.review, k.total)} of records`);
    set('avgAge', k.avgAge === null ? '–' : k.avgAge.toFixed(1), k.avgAge === null ? 'No age data' : `days · oldest ${fmtNum(k.maxAge)} days`);
    set('aged', k.aged, `Total Age ≥ ${cfg.agedFrom} days · ${pct(k.aged, k.total)}`);
    set('assignees', k.assignees, `${fmtNum(k.notPended)} record(s) without a pended date`);
  }

  /* ---------- report table ---------- */
  const R = {
    text: (d, t) => (t === 'display' ? esc(d) : d ?? ''),
    long: (d, t) => (t === 'display' ? `<span class="trunc" title="${esc(d)}">${esc(d)}</span>` : d ?? ''),
    num: (d, t) => (t === 'display' || t === 'filter' ? (d === null ? '' : fmtNum(d)) : d ?? -1),
    status: (d, t) => (t === 'display' ? `<span class="badge badge-${/review/i.test(d) ? 'other' : d ? 'pending' : 'blank'}">${esc(d || 'Blank')}</span>` : d),
    pended: (d, t, row) => {
      if (t === 'sort' || t === 'type') return row.pendedAt ? row.pendedAt.getTime() : 0;
      if (!row.pendedAt) return '';
      const time = row.pendedAt.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
      return `${fmtDate(row.pendedDate)} ${time}`;
    }
  };

  function renderTable(rows) {
    $('#iqTableCount').textContent = `${fmtNum(rows.length)} record${rows.length === 1 ? '' : 's'}`;
    if (dt) { dt.clear().rows.add(rows).draw(false); return; }
    const { exportButtons, language } = APEX.tables.helpers;
    dt = new DataTable('#iqTable', {
      data: rows,
      columns: [
        { data: 'inquiry', title: 'Inquiry Number', render: R.text },
        { data: 'dcn', title: 'DCN', render: R.text },
        { data: 'status', title: 'Status', render: R.status },
        { data: 'subStatus', title: 'Sub-status', render: R.text },
        { data: 'assignedTo', title: 'Assigned To', render: R.text },
        { data: 'operator', title: 'Last Updated Operator ID', render: R.text },
        { data: 'totalAge', title: 'Total Age', type: 'num', className: 'dt-center', render: R.num },
        { data: 'ageBucket', title: 'Age Bucket', render: R.text },
        { data: 'pendedAt', title: 'Pended Date', type: 'num', render: R.pended },
        { data: 'lastNote', title: 'Last Note', render: R.long }
      ].map(c => ({ ...c, defaultContent: '' })),
      order: [[6, 'desc']],
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100, { label: 'All', value: -1 }],
      autoWidth: false,
      layout: { topStart: ['pageLength', 'buttons'], topEnd: null, bottomStart: 'info', bottomEnd: 'paging' },
      language: { ...language, emptyTable: 'No IQT records match the current search and filters' },
      buttons: exportButtons('IQT Detailed Report', { pageSize: 'A3' })
    });
  }

  /* ---------- render pipeline ---------- */
  function render() {
    if (!state.all) return;
    try {
      const global = D.applySearch(D.applySelections(state.all, selections(gms)), state.gQuery);
      state.model = D.build(global);
      renderKpis(state.model.kpi);
      IC.render(state.model);

      const searchNote = state.gQuery ? [`<span class="filter-chip">Search: “${esc(state.gQuery)}”<button type="button" data-search="1" aria-label="Clear search">&times;</button></span>`] : [];
      chips('#iqFilterChips', gms, cfg.globalFilters, searchNote);
      $('#iqSearchCount').textContent = global.length === state.all.length ? `${fmtNum(state.all.length)} records` : `${fmtNum(global.length)} of ${fmtNum(state.all.length)} records`;

      const topts = Object.fromEntries(cfg.tableFilters.map(f => [f.key, D.options(global, f.key)]));
      Object.entries(tms).forEach(([k, m]) => m.setOptions(topts[k.slice(1)] || []));
      const tsel = Object.fromEntries(Object.entries(selections(tms)).map(([k, s]) => [k.slice(1), s]));
      const tableRows = D.applyDateRange(D.applySearch(D.applySelections(global, tsel), state.tQuery), state.from, state.to);
      renderTable(tableRows);
      const tdefs = cfg.tableFilters.map(f => ({ key: `t${f.key}`, label: f.label }));
      const range = state.from || state.to ? [`<span class="filter-chip">Pended: ${state.from ? fmtDate(state.from) : '…'} – ${state.to ? fmtDate(state.to) : '…'}<button type="button" data-range="1" aria-label="Clear date range">&times;</button></span>`] : [];
      chips('#iqTableChips', tms, tdefs, range);
    } catch (err) {
      console.error('[APEX IQT] render failed', err);
      toast(`IQT dashboard could not be rendered: ${err.message}`, 'error', 7000);
    }
  }

  /* ---------- data events ---------- */
  source.on('data', ({ records, sheet, at, file }) => {
    state.all = records;
    state.last = at;
    cfg.globalFilters.forEach(f => gms[f.key] && gms[f.key].setOptions(D.options(records, f.key)));
    $('#iqState').hidden = true;
    $('#iqContent').hidden = false;
    $('#iqSourceInfo').textContent = `${file} › ${sheet || cfg.sheet} · ${fmtNum(records.length)} records`;
    render();
    if (state.loaded) toast(`IQT workbook updated · ${fmtNum(records.length)} records refreshed`, 'success');
    state.loaded = true;
    liveStatus();
  });
  source.on('status', () => { $('#iqRefresh').classList.remove('spinning'); liveStatus(); });
  source.on('error', err => {
    console.error('[APEX IQT] data error', err);
    $('#iqRefresh').classList.remove('spinning');
    const msg = err && err.message ? err.message : String(err);
    if (!state.loaded) showState('IQT Data Unavailable', msg, true);
    else { setStatus('warn', `Last good data ${fmtTime(state.last)} · retrying`); toast(`Could not read IQT workbook (will retry): ${msg}`, 'error', 6000); }
  });
  source.on('needsFile', ({ reconnect }) => showState('Connect IQT Workbook',
    reconnect ? 'Click below to re-authorise read access to the IQT workbook.' : `Connect "${cfg.files[0]}" once to enable the IQT dashboards. The file is opened read-only and is never modified.`, true));

  /* ---------- wiring ---------- */
  function ensureStarted() {
    if (state.started) return;
    state.started = true;
    source.start().catch(err => showState('IQT Data Unavailable', err.message || String(err), true));
  }

  $$('[data-ifilter]').forEach(el => {
    (el.closest('#iqTableFilters') ? tms : gms)[el.dataset.ifilter] = filters.createMultiSelect(el, render);
  });

  const gSearch = $('#iqSearch');
  const onGlobalSearch = debounce(() => { state.gQuery = gSearch.value.trim(); $('#iqSearchClear').hidden = !gSearch.value; render(); }, 120);
  gSearch.addEventListener('input', onGlobalSearch);
  $('#iqSearchClear').addEventListener('click', () => { gSearch.value = ''; state.gQuery = ''; $('#iqSearchClear').hidden = true; render(); gSearch.focus(); });

  const tSearch = $('#iqTableSearch');
  tSearch.addEventListener('input', debounce(() => { state.tQuery = tSearch.value.trim(); render(); }, 120));
  $('#iqDateFrom').addEventListener('change', e => { state.from = parseInputDate(e.target.value); render(); });
  $('#iqDateTo').addEventListener('change', e => { state.to = parseInputDate(e.target.value); render(); });

  $('#iqResetFilters').addEventListener('click', () => {
    Object.values(gms).forEach(m => m.clear());
    gSearch.value = ''; state.gQuery = ''; $('#iqSearchClear').hidden = true;
    render();
  });
  $('#iqTableReset').addEventListener('click', () => {
    Object.values(tms).forEach(m => m.clear());
    tSearch.value = ''; state.tQuery = '';
    $('#iqDateFrom').value = ''; $('#iqDateTo').value = ''; state.from = state.to = null;
    render();
  });
  $('#iqFilterChips').addEventListener('click', e => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.search) { gSearch.value = ''; state.gQuery = ''; $('#iqSearchClear').hidden = true; }
    else gms[b.dataset.key].remove(b.dataset.val);
    render();
  });
  $('#iqTableChips').addEventListener('click', e => {
    const b = e.target.closest('button');
    if (!b) return;
    if (b.dataset.range) { $('#iqDateFrom').value = ''; $('#iqDateTo').value = ''; state.from = state.to = null; }
    else tms[b.dataset.key].remove(b.dataset.val);
    render();
  });

  $('#iqRefresh').addEventListener('click', () => {
    if (source.getMode() && source.getMode() !== 'manual') { $('#iqRefresh').classList.add('spinning'); source.refresh(); }
    else source.connect();
  });
  $('#iqConnect').addEventListener('click', () => {
    showState('Loading IQT Data', 'Reading the IQT workbook…');
    source.connect().then(r => { if (r === 'aborted' && !state.loaded) { state.started = false; ensureStarted(); } });
  });
  $('#iqRetry').addEventListener('click', () => { showState('Loading IQT Data', 'Reading the IQT workbook…'); state.started = false; ensureStarted(); });

  IC.onPick((key, value) => {
    if (!gms[key]) return;
    gms[key].set([value]);
    render();
    toast(`IQT filtered: ${(cfg.globalFilters.find(f => f.key === key) || {}).label || key} = ${value}`, 'info', 3000);
  });
  theme.onChange(() => { if (state.model) IC.render(state.model); });

  $$('.tab').forEach(t => t.addEventListener('click', () => { if (t.dataset.tab === 'iqt') ensureStarted(); }));
  if ($('#tab-iqt').classList.contains('active')) ensureStarted();
})();
