/* APEX Dashboard - PREPAY › JIRA Analytics (source: "FOR JIRA" sheet of the PREPAY workbook) */
window.APEX = window.APEX || {};

APEX.prepayJira = (() => {
  const { $, $$, str, esc, toDate, fmtDate, fmtNum, animateCount, toast } = APEX.utils;
  const B = APEX.prepayConfig.blankLabel;
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const COLUMNS = {
    domain: ['Domain'],
    associate: ['Associate'],
    issue: ['Issue'],
    dcn: ['DCN'],
    hcid: ['HCID'],
    accums: ['ACCUMS CODE', 'Accums Code'],
    sccf: ['SCCF#', 'SCCF'],
    claimType: ['Claim type', 'Claim Type'],
    dateEntered: ['Date entered', 'Date Entered'],
    dateCreated: ['Date JIRA Created', 'Date JIRA created'],
    prepay2: ['PREPAY2', 'PREPAY2 Ticket Number'],
    status: ['Status'],
    notes: ['Notes'],
    assignedTo: ['Assigned to', 'Assigned To']
  };

  const FILTERS = [
    { key: 'associate', label: 'Associate' },
    { key: 'status', label: 'Status' },
    { key: 'assignedTo', label: 'Assigned To' },
    { key: 'claimType', label: 'Claim Type' },
    { key: 'enteredMonth', label: 'Date Entered' },
    { key: 'createdMonth', label: 'Date JIRA Created' }
  ];

  const ms = {};
  let all = null;
  let dt = null;

  const monthLabel = d => (d ? `${MONTHS[d.getMonth()]} ${d.getFullYear()}` : '');
  const clean = v => { const s = str(v); return s === '0' ? '' : s; };
  /* "Assigned to A, Praveen - Praveen.A@anthem.com" -> "A, Praveen" */
  const cleanAssignee = v => clean(v).replace(/^(assigned\s+to|assignee)\s*:?\s*/i, '').replace(/\S+@\S+/g, '').replace(/[\s\-–]+$/, '').trim();

  /* ---------- parse ---------- */
  function parse(ws) {
    if (!ws) return [];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
    const hi = rows.findIndex(r => r.some(c => str(c).toLowerCase() === 'associate'));
    if (hi < 0) return [];
    const lookup = new Map(rows[hi].map((h, i) => [str(h).toLowerCase(), i]));
    const idx = {};
    Object.entries(COLUMNS).forEach(([k, aliases]) => {
      const a = aliases.find(x => lookup.has(x.toLowerCase()));
      if (a) idx[k] = lookup.get(a.toLowerCase());
    });

    return rows.slice(hi + 1).map(row => {
      const r = {};
      Object.entries(idx).forEach(([k, i]) => { r[k] = row[i]; });
      ['dateEntered', 'dateCreated'].forEach(k => {
        const raw = r[k];
        r[k] = toDate(raw);
        r[`${k}Text`] = r[k] ? '' : clean(raw);
      });
      ['domain', 'associate', 'issue', 'dcn', 'hcid', 'accums', 'sccf', 'claimType', 'prepay2', 'status', 'notes'].forEach(k => { r[k] = clean(r[k]); });
      r.assignedRaw = clean(r.assignedTo);
      r.assignedTo = cleanAssignee(r.assignedTo);
      r.enteredMonth = monthLabel(r.dateEntered);
      r.createdMonth = monthLabel(r.dateCreated);
      return r;
    }).filter(r => r.associate || r.dcn || r.issue || r.prepay2);
  }

  /* ---------- model ---------- */
  const fv = (r, k) => r[k] || B;
  const lc = v => (v || '').toLowerCase();

  function countBy(recs, key) {
    const m = new Map();
    recs.forEach(r => m.set(fv(r, key), (m.get(fv(r, key)) || 0) + 1));
    return [...m.entries()].map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
  }

  function options(recs, key) {
    const list = countBy(recs, key);
    if (key === 'enteredMonth' || key === 'createdMonth') {
      return list.sort((a, b) => (a.name === B) - (b.name === B) || new Date(`1 ${a.name}`) - new Date(`1 ${b.name}`));
    }
    return list.sort((a, b) => (a.name === B) - (b.name === B) || a.name.localeCompare(b.name));
  }

  function filtered() {
    const active = Object.entries(ms).map(([k, m]) => [k, m.values()]).filter(([, s]) => s.size);
    return active.length ? all.filter(r => active.every(([k, s]) => s.has(fv(r, k)))) : all;
  }

  /* ---------- KPIs ---------- */
  function renderKpis(recs) {
    const total = recs.length;
    const closed = recs.filter(r => lc(r.status) === 'closed').length;
    const resolved = recs.filter(r => lc(r.status) === 'resolved').length;
    const open = recs.filter(r => !r.status || lc(r.status) === 'open').length;
    const other = total - closed - resolved - open;
    const associates = new Set(recs.map(r => r.associate).filter(Boolean)).size;
    const pct = n => (total ? `${((n / total) * 100).toFixed(1)}%` : '0%');

    const set = (k, v, sub) => { animateCount($(`[data-jkpi="${k}"]`), v); $(`[data-jkpi-sub="${k}"]`).textContent = sub; };
    set('total', total, other ? `${fmtNum(other)} invalid / duplicate / other` : 'All tickets in FOR JIRA');
    set('closed', closed, `${pct(closed)} of tickets`);
    set('resolved', resolved, `${pct(resolved)} of tickets`);
    set('open', open, `${pct(open)} · status Open or blank`);
    set('associates', associates, associates ? `${(total / associates).toFixed(1)} tickets per associate` : '–');
  }

  /* ---------- charts ---------- */
  function doughnut(id, rows, c, pick) {
    const PC = APEX.prepayCharts;
    const H = APEX.charts.helpers;
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
          borderColor: c.surface,
          borderWidth: 3,
          hoverOffset: 12,
          datalabels: {
            display: ctx => total && ctx.dataset.data[ctx.dataIndex] / total >= 0.04,
            color: '#fff',
            font: { weight: '700', size: 11.5 },
            textShadowColor: 'rgba(0,0,0,.5)',
            textShadowBlur: 4,
            formatter: v => `${fmtNum(v)}\n${((v / total) * 100).toFixed(0)}%`,
            textAlign: 'center'
          }
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '58%',
        animation: { animateRotate: true, duration: 1200, easing: 'easeOutQuart' },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length) pick(rows[els[0].index].name); },
        plugins: {
          legend: { position: 'right', labels: { color: c.text, usePointStyle: true, boxWidth: 9, padding: 14, font: { weight: '600' } } },
          tooltip: {
            ...H.tooltip(c),
            callbacks: {
              label: item => ` ${item.label}: ${fmtNum(item.raw)} (${total ? ((item.raw / total) * 100).toFixed(1) : 0}%)`,
              footer: () => 'Click to filter'
            },
            footerColor: c.muted,
            footerFont: { size: 11, style: 'italic', weight: '400' }
          }
        }
      }
    });
  }

  function pickFilter(key) {
    return value => {
      ms[key].set([value]);
      render();
      toast(`JIRA filtered: ${FILTERS.find(f => f.key === key).label} = ${value}`, 'info', 3000);
    };
  }

  function renderCharts(recs) {
    if (!window.Chart) return;
    const c = APEX.theme.colors();
    const PC = APEX.prepayCharts;
    const byAssoc = countBy(recs, 'associate');
    const byAssigned = countBy(recs, 'assignedTo');
    doughnut('jiraStatusChart', countBy(recs, 'status'), c, pickFilter('status'));
    doughnut('jiraClaimChart', countBy(recs, 'claimType'), c, pickFilter('claimType'));
    PC.bar('jiraAssocChart', byAssoc, c, { label: 'Tickets', onPick: pickFilter('associate') });
    PC.bar('jiraAssignedChart', byAssigned, c, { label: 'Tickets', alt: true, onPick: pickFilter('assignedTo') });
    $('#jiraAssocBadge').textContent = `${byAssoc.length} associates`;
    $('#jiraAssignedBadge').textContent = `${byAssigned.length} assignees`;
  }

  /* ---------- table ---------- */
  const R = {
    text: (d, t) => (t === 'display' ? esc(d) : d ?? ''),
    long: (d, t) => (t === 'display' ? `<span class="trunc" title="${esc(d)}">${esc(d)}</span>` : d ?? ''),
    status: (d, t) => {
      if (t !== 'display') return d || '';
      const s = lc(d);
      const cls = s === 'resolved' || s === 'closed' ? 'completed' : s === 'open' ? 'pending' : !s ? 'blank' : 'other';
      return `<span class="badge badge-${cls}">${esc(d || 'Blank')}</span>`;
    },
    date: textKey => (d, t, row) => {
      if (t === 'sort' || t === 'type') return d instanceof Date ? d.getTime() : 0;
      return d instanceof Date ? fmtDate(d) : t === 'display' ? esc(row[textKey]) : row[textKey] || '';
    }
  };

  const TABLE_COLUMNS = [
    { data: 'associate', title: 'Associate', render: R.text },
    { data: 'dcn', title: 'DCN', render: R.text },
    { data: 'hcid', title: 'HCID', render: R.text },
    { data: 'issue', title: 'Issue', render: R.long },
    { data: 'accums', title: 'ACCUMS CODE', render: R.text },
    { data: 'claimType', title: 'Claim Type', render: R.text },
    { data: 'dateEntered', title: 'Date Entered', type: 'num', render: R.date('dateEnteredText') },
    { data: 'dateCreated', title: 'Date JIRA Created', type: 'num', render: R.date('dateCreatedText') },
    { data: 'prepay2', title: 'PREPAY2 Ticket Number', render: R.text },
    { data: 'status', title: 'Status', render: R.status },
    { data: 'assignedTo', title: 'Assigned To', render: R.text },
    { data: 'notes', title: 'Notes', render: R.long }
  ];

  function renderTable(recs) {
    $('#jiraTableCount').textContent = recs.length === all.length ? `${fmtNum(recs.length)} tickets` : `${fmtNum(recs.length)} of ${fmtNum(all.length)} tickets`;
    if (dt) { dt.clear().rows.add(recs).draw(false); return; }
    const { exportButtons, layout, language } = APEX.tables.helpers;
    dt = new DataTable('#jiraTable', {
      data: recs,
      columns: TABLE_COLUMNS.map(c => ({ ...c, defaultContent: '' })),
      order: [[6, 'desc']],
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100, { label: 'All', value: -1 }],
      deferRender: true,
      autoWidth: false,
      layout,
      language: { ...language, searchPlaceholder: 'Search JIRA tickets…', emptyTable: 'No JIRA tickets match the selected filters' },
      buttons: exportButtons('PREPAY JIRA Report', { pageSize: 'A3' })
    });
  }

  function renderChips() {
    $('#jiraFilterChips').innerHTML = FILTERS.flatMap(f => [...ms[f.key].values()].map(v =>
      `<span class="filter-chip">${esc(f.label)}: ${esc(v)}<button type="button" data-key="${f.key}" data-val="${esc(v)}" aria-label="Remove filter">&times;</button></span>`)).join('');
  }

  /* ---------- render ---------- */
  function render() {
    if (!all) return;
    try {
      const recs = filtered();
      renderKpis(recs);
      renderCharts(recs);
      renderTable(recs);
      renderChips();
    } catch (err) {
      console.error('[APEX JIRA] render failed', err);
      toast(`JIRA Analytics could not be rendered: ${err.message}`, 'error', 7000);
    }
  }

  function load(records, sheetName) {
    const empty = !records || !records.length;
    $('#jiraEmpty').hidden = !empty;
    $('#jiraBody').hidden = empty;
    $('#jiraSourceInfo').textContent = sheetName
      ? `Source: "${sheetName}" sheet · ${fmtNum(records ? records.length : 0)} tickets`
      : `"FOR JIRA" sheet not found in the PREPAY workbook`;
    if (empty) { all = null; return; }
    all = records;
    FILTERS.forEach(f => ms[f.key].setOptions(options(all, f.key)));
    render();
  }

  function init() {
    $$('[data-jfilter]').forEach(el => { ms[el.dataset.jfilter] = APEX.filters.createMultiSelect(el, render); });
    $('#jiraResetFilters').addEventListener('click', () => { Object.values(ms).forEach(m => m.clear()); render(); });
    $('#jiraFilterChips').addEventListener('click', e => {
      const b = e.target.closest('button[data-key]');
      if (!b) return;
      ms[b.dataset.key].remove(b.dataset.val);
      render();
    });
    APEX.theme.onChange(() => { if (all) renderCharts(filtered()); });
  }

  return { parse, load, init };
})();
