/* APEX Dashboard - PREPAY reports (DataTables) */
window.APEX = window.APEX || {};

APEX.prepayReports = (() => {
  const { $, $$, esc, fmtNum, fmtDate } = APEX.utils;
  const cfg = APEX.prepayConfig;
  const D = APEX.prepayData;
  const B = cfg.blankLabel;
  let dt = null;
  let active = null;
  let records = [];

  /* ---------- cell renderers ---------- */
  const R = {
    text: (d, t) => (t === 'display' ? esc(d) : d ?? ''),
    long: (d, t) => (t === 'display' ? `<span class="trunc" title="${esc(d)}">${esc(d)}</span>` : d ?? ''),
    num: (d, t) => (t === 'display' || t === 'filter' || t === 'export' ? (d === null || d === undefined || d === '' ? '' : fmtNum(d)) : d ?? -1),
    dec: (d, t) => (t === 'display' || t === 'filter' ? (d === null || d === undefined ? '' : Number(d).toFixed(1)) : d ?? -1),
    pct: (d, t) => (t === 'display' || t === 'filter' ? `${Number(d || 0).toFixed(1)}%` : d ?? 0),
    date: (d, t) => (t === 'sort' || t === 'type' ? (d instanceof Date ? d.getTime() : 0) : d instanceof Date ? fmtDate(d) : ''),
    status: (d, t) => (t === 'display' ? `<span class="badge badge-${d === 'Finalized' ? 'completed' : 'pending'}">${esc(d)}</span>` : d),
    verified: (d, t) => {
      if (t !== 'display') return d || '';
      const v = (d || '').toLowerCase();
      return d ? `<span class="badge badge-${v === 'valid' ? 'completed' : v === 'invalid' ? 'other' : 'blank'}">${esc(d)}</span>` : '';
    },
    urgency: (d, t) => (t === 'display' ? (d === null ? '' : `<span class="cat-chip ${d >= cfg.highUrgency ? 'cat-pends' : ''}">${d}</span>`) : d ?? -1)
  };

  const col = (data, title, type = 'text') => ({ data, title, type, render: R[type], className: ['num', 'dec', 'pct', 'urgency'].includes(type) ? 'dt-center' : '' });

  const recordCols = {
    snapshot: col('snapshotDate', 'Snapshot', 'date'),
    dcn: col('dcn', 'DCN'),
    assigned: col('assignedTo', 'Assigned To'),
    sub: col('subStatus', 'Pega Sub-status'),
    urgency: col('urgency', 'Urgency', 'urgency'),
    skill: col('skill', 'Claim Skill'),
    updates: col('updates', 'Updates'),
    status: col('status', 'Status', 'status'),
    verified: col('verified', 'Verified', 'verified'),
    wgs: col('wgs', 'WGS Status Code'),
    dm: col('dmAge', 'DM Age', 'num'),
    dmBucket: col('dmBucket', 'DM Age Bucket'),
    touched: col('lastTouched', 'Days Since Touched', 'num'),
    touchBucket: col('touchBucket', 'Touch Bucket'),
    research: col('lastResearch', 'Last Research Note', 'date'),
    algorithm: col('algorithm', 'Algorithm'),
    ticket: col('ticket', 'Ticket #'),
    note: col('lastNote', 'Last Note', 'long'),
    comment: col('comment', 'Comment', 'long')
  };

  const summaryCols = name => [
    col('name', name),
    col('total', 'Total', 'num'),
    col('finalized', 'Finalized', 'num'),
    col('open', 'Open', 'num'),
    col('finalizedPct', 'Finalized %', 'pct'),
    col('valid', 'Valid', 'num'),
    col('invalid', 'Invalid', 'num'),
    col('high', 'High Urgency', 'num'),
    col('avgDm', 'Avg DM Age', 'dec'),
    col('maxDm', 'Max DM Age', 'num')
  ];

  const c = recordCols;

  /* ---------- report definitions ---------- */
  const REPORTS = [
    {
      id: 'detail', title: 'Detailed Records',
      desc: 'Every record from all October snapshot sheets, combined into one dataset.',
      build: recs => ({ rows: recs, columns: [c.snapshot, c.dcn, c.assigned, c.sub, c.urgency, c.skill, c.updates, c.status, c.verified, c.wgs, c.dm, c.touched, c.research, c.algorithm, c.ticket, c.note, c.comment], order: [[0, 'desc'], [10, 'desc']] })
    },
    {
      id: 'aging', title: 'Aging Report',
      desc: 'DM Age distribution per Assigned To, with average and maximum age.',
      build: recs => {
        const buckets = [...cfg.dmBuckets.map(b => b[2]), 'Not recorded'];
        const groups = D.groupSummary(recs, r => r.assignedTo || B);
        const used = buckets.filter(b => recs.some(r => r.dmBucket === b));
        const rows = groups.map(g => {
          const o = { name: g.name, total: g.total, avgDm: g.avgDm, maxDm: g.maxDm };
          used.forEach((b, i) => { o[`b${i}`] = g.rows.filter(r => r.dmBucket === b).length; });
          return o;
        });
        return {
          rows,
          columns: [col('name', 'Assigned To'), ...used.map((b, i) => col(`b${i}`, b, 'num')), col('total', 'Total', 'num'), col('avgDm', 'Avg DM Age', 'dec'), col('maxDm', 'Max DM Age', 'num')],
          order: [[used.length + 2, 'desc']]
        };
      }
    },
    {
      id: 'high', title: 'High Urgency Report',
      desc: `Records with Urgency ≥ ${cfg.highUrgency}, highest urgency and oldest first.`,
      build: recs => ({ rows: recs.filter(r => r.high), columns: [c.snapshot, c.dcn, c.urgency, c.assigned, c.sub, c.skill, c.status, c.updates, c.dm, c.touched, c.wgs, c.note], order: [[2, 'desc'], [8, 'desc']] })
    },
    {
      id: 'followup', title: 'Follow-Up Required',
      desc: 'Open (not finalized) records that still need action, longest untouched first.',
      build: recs => ({
        rows: recs.filter(r => !r.finalized).map(r => Object.assign(Object.create(r), { reason: r.updates || 'No update recorded' })),
        columns: [c.snapshot, c.dcn, c.assigned, col('reason', 'Follow-Up Reason'), c.sub, c.urgency, c.dm, c.touched, c.research, c.ticket, c.note, c.comment],
        order: [[7, 'desc'], [6, 'desc']]
      })
    },
    {
      id: 'touched', title: 'Last Touched Report',
      desc: 'Days since each record was last touched in CIW, with the last research note date.',
      build: recs => ({ rows: recs, columns: [c.snapshot, c.dcn, c.assigned, c.touched, c.touchBucket, c.research, c.status, c.updates, c.sub, c.note], order: [[3, 'desc']] })
    },
    {
      id: 'assigned', title: 'Assigned To Report',
      desc: 'Productivity summary per associate.',
      build: recs => ({ rows: D.groupSummary(recs, r => r.assignedTo || B), columns: summaryCols('Assigned To'), order: [[1, 'desc']] })
    },
    {
      id: 'substatus', title: 'Sub-status Report',
      desc: 'Summary per Pega sub-status.',
      build: recs => ({ rows: D.groupSummary(recs, r => r.subStatus || B), columns: summaryCols('Pega Sub-status'), order: [[1, 'desc']] })
    },
    {
      id: 'wgs', title: 'WGS Status Report',
      desc: 'Summary per individual WGS status code (records with several codes count toward each code).',
      build: recs => ({ rows: D.groupSummary(recs, r => (r.wgsCodes.length ? r.wgsCodes : [B])), columns: summaryCols('WGS Status Code'), order: [[1, 'desc']] })
    }
  ];

  function drawTable() {
    const def = REPORTS.find(r => r.id === active) || REPORTS[0];
    const { rows, columns, order } = def.build(records);
    const table = $('#ppReportTable');
    const { exportButtons, layout, language } = APEX.tables.helpers;

    if (dt) { dt.destroy(); dt = null; }
    table.innerHTML = '';

    $('#ppReportTitle').textContent = def.title;
    $('#ppReportDesc').textContent = def.desc;
    $('#ppReportCount').textContent = `${fmtNum(rows.length)} row${rows.length === 1 ? '' : 's'}`;

    dt = new DataTable(table, {
      data: rows,
      columns: columns.map(x => ({ ...x, defaultContent: '', type: ['num', 'dec', 'pct', 'urgency', 'date'].includes(x.type) ? 'num' : undefined })),
      order,
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100, { label: 'All', value: -1 }],
      deferRender: true,
      autoWidth: false,
      layout,
      language: { ...language, emptyTable: 'No records match the selected PREPAY filters' },
      buttons: exportButtons(`PREPAY ${def.title}`, { pageSize: columns.length > 10 ? 'A3' : 'A4' })
    });
  }

  function select(id) {
    active = id;
    try { localStorage.setItem(cfg.storageKeys.report, id); } catch { /* ignore */ }
    $$('.pp-pill').forEach(p => {
      const on = p.dataset.report === id;
      p.classList.toggle('active', on);
      p.setAttribute('aria-selected', on);
    });
    if (window.DataTable) drawTable();
  }

  function init() {
    const bar = $('#ppReportTabs');
    bar.innerHTML = REPORTS.map((r, i) => `<button type="button" class="pp-pill" role="tab" data-report="${r.id}"><b>${i + 1}</b>${esc(r.title)}</button>`).join('');
    bar.addEventListener('click', e => {
      const b = e.target.closest('.pp-pill');
      if (b) select(b.dataset.report);
    });
    let saved = null;
    try { saved = localStorage.getItem(cfg.storageKeys.report); } catch { /* ignore */ }
    active = REPORTS.some(r => r.id === saved) ? saved : REPORTS[0].id;
  }

  function render(recs) {
    records = recs;
    select(active);
  }

  return { init, render };
})();
