/* APEX Dashboard - NON COB detailed reports (DataTables) */
window.APEX = window.APEX || {};

APEX.noncobReports = (() => {
  const { $, $$, esc, fmtNum, fmtDate } = APEX.utils;
  const cfg = APEX.noncobConfig;
  let dt = null;
  let active = 'ageing';
  let data = { ageing: [], rstatus: [], opid: [] };

  const R = {
    text: (d, t) => (t === 'display' ? esc(d) : d ?? ''),
    long: (d, t) => (t === 'display' ? `<span class="trunc" title="${esc(d)}">${esc(d)}</span>` : d ?? ''),
    num: (d, t) => (t === 'display' || t === 'filter' ? (d === null || d === undefined ? '' : fmtNum(d)) : d ?? -1),
    money: (d, t) => (t === 'display' || t === 'filter' ? (d === null || d === undefined ? '' : Number(d).toLocaleString('en-US', { style: 'currency', currency: 'USD' })) : d ?? -1),
    date: (d, t) => (t === 'sort' || t === 'type' ? (d instanceof Date ? d.getTime() : 0) : d instanceof Date ? fmtDate(d) : ''),
    carry: (d, t) => (t === 'display' ? `<span class="badge badge-${d === 'Carry Over' ? 'other' : 'blank'}">${esc(d)}</span>` : d)
  };
  const col = (data, title, type = 'text') => ({ data, title, render: R[type], type: ['num', 'money', 'date'].includes(type) ? 'num' : undefined, className: ['num', 'money'].includes(type) ? 'dt-center' : '' });

  const REPORTS = [
    {
      id: 'ageing', title: 'Ageing Report', sheet: 'ageing',
      desc: 'All records from the Ageing sheet with current and pend age.',
      columns: [
        col('weekEnding', 'Week Ending', 'date'), col('associate', 'Associate'), col('team', 'Team'), col('manager', 'Manager ID'),
        col('claimId', 'Claim ID'), col('memberId', 'Member ID'), col('errNum', 'Err Num'), col('status', 'Status'),
        col('pendDesc', 'Pend Code Desc'), col('recoveryType', 'Recovery Type'), col('currAge', 'CURR_AGE', 'num'), col('currAgeCat', 'CURR_AGE_CAT'),
        col('pendAge', 'PEND_AGE_CURRENT', 'num'), col('pendAgeCat', 'PEND_AGE_CURRENT_CAT'), col('state', 'State'), col('mbu', 'MBU'),
        col('provName', 'Provider', 'long'), col('grpName', 'Group', 'long'), col('paid', 'Paid Amount', 'money'), col('project', 'Project', 'long')
      ],
      order: [[10, 'desc']]
    },
    {
      id: 'rstatus', title: 'R Status Report', sheet: 'rstatus',
      desc: 'All records from the R Status sheet.',
      columns: [
        col('weekEnding', 'Week Ending', 'date'), col('associate', 'Associate'), col('team', 'Team'), col('manager', 'Manager'),
        col('rStatus', 'R Status'), col('claimId', 'Claim ID'), col('errNum', 'Err Num'), col('rStatusDate', 'R Status Date', 'date'),
        col('recoveryType', 'Recovery Type'), col('adjReason', 'Adjustment Reason', 'long'), col('srcAdjReason', 'Src Adjustment Reason', 'long'),
        col('state', 'State'), col('mbu', 'MBU'), col('paid', 'Paid Amount', 'money'), col('grpName', 'Group', 'long'), col('provName', 'Provider', 'long')
      ],
      order: [[7, 'desc']]
    },
    {
      id: 'opid', title: 'OPID Report', sheet: 'opid',
      desc: 'All records from the OPID sheet with carry-over status and age.',
      columns: [
        col('weekEnding', 'Week Ending', 'date'), col('associate', 'Associate'), col('team', 'Team'), col('processor', 'Processor'),
        col('manager', 'Manager'), col('mktManager', 'Market Manager'), col('carryOver', 'Carry Over', 'carry'), col('claimNumber', 'Claim Number'),
        col('ageDays', 'Age (Days)', 'num'), col('ageBucket', 'Age Bucket'), col('pendException', 'Pend Exception'), col('queueStatus', 'Queue Status'),
        col('pendCode', 'Pend Code', 'long'), col('lastProcessed', 'Date Last Processed', 'date'), col('state', 'State'), col('mbu', 'MBU'), col('area', 'Area')
      ],
      order: [[8, 'desc']]
    }
  ];

  function draw() {
    const def = REPORTS.find(r => r.id === active) || REPORTS[0];
    const rows = data[def.sheet] || [];
    const table = $('#ncReportTable');
    const { exportButtons, layout, language } = APEX.tables.helpers;
    if (dt) { dt.destroy(); dt = null; }
    table.innerHTML = '';
    $('#ncReportTitle').textContent = def.title;
    $('#ncReportDesc').textContent = def.desc;
    $('#ncReportCount').textContent = `${fmtNum(rows.length)} row${rows.length === 1 ? '' : 's'}`;
    dt = new DataTable(table, {
      data: rows,
      columns: def.columns.map(c => ({ ...c, defaultContent: '' })),
      order: def.order,
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100, { label: 'All', value: -1 }],
      deferRender: true,
      autoWidth: false,
      layout,
      language: { ...language, emptyTable: 'No records match the selected NON COB filters' },
      buttons: exportButtons(`NON COB ${def.title}`, { pageSize: 'A3' })
    });
  }

  function select(id) {
    active = id;
    try { localStorage.setItem(cfg.storageKeys.report, id); } catch { /* ignore */ }
    $$('#ncReportTabs .pp-pill').forEach(p => {
      const on = p.dataset.report === id;
      p.classList.toggle('active', on);
      p.setAttribute('aria-selected', on);
    });
    if (window.DataTable) draw();
  }

  function init() {
    const bar = $('#ncReportTabs');
    bar.innerHTML = REPORTS.map((r, i) => `<button type="button" class="pp-pill" role="tab" data-report="${r.id}"><b>${i + 1}</b>${esc(r.title)}</button>`).join('');
    bar.addEventListener('click', e => { const b = e.target.closest('.pp-pill'); if (b) select(b.dataset.report); });
    let saved = null;
    try { saved = localStorage.getItem(cfg.storageKeys.report); } catch { /* ignore */ }
    if (REPORTS.some(r => r.id === saved)) active = saved;
  }

  function render(filtered) {
    data = filtered;
    select(active);
  }

  return { init, render };
})();
