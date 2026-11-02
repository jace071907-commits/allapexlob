/* APEX Dashboard - DataTables (Container #6 matrix, Container #7 detail records) */
window.APEX = window.APEX || {};

APEX.tables = (() => {
  const { $, esc, fmtNum, fmtDate, slug } = APEX.utils;
  const { detailColumns, blankLabel } = APEX.config;
  let matrixDT = null;
  let detailDT = null;

  const ICON = {
    excel: '<svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M9 12l6 6M15 12l-6 6"/></svg>',
    csv: '<svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h8"/></svg>',
    pdf: '<svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M9 15h1.5a1.5 1.5 0 0 0 0-3H9v6"/></svg>',
    print: '<svg viewBox="0 0 24 24"><path d="M6 9V2h12v7"/><rect x="2" y="9" width="20" height="9" rx="2"/><path d="M6 14h12v8H6z"/></svg>'
  };

  const stamp = () => {
    const d = new Date();
    return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}`;
  };

  function exportButtons(title, { footer = false, pageSize = 'A4' } = {}) {
    const base = {
      title: `APEX Dashboard - ${title}`,
      filename: () => `APEX_${title.replace(/[^a-z0-9]+/gi, '_')}_${stamp()}`,
      footer,
      exportOptions: { columns: ':visible' }
    };
    const accent = APEX.theme.colors().accent || '#2f8cff';
    return [
      { extend: 'excelHtml5', text: `${ICON.excel}<span>Excel</span>`, ...base },
      { extend: 'csvHtml5', text: `${ICON.csv}<span>CSV</span>`, ...base },
      {
        extend: 'pdfHtml5', text: `${ICON.pdf}<span>PDF</span>`, orientation: 'landscape', pageSize, ...base,
        customize: doc => {
          doc.defaultStyle.fontSize = 7.5;
          doc.styles.tableHeader.fillColor = APEX.theme.colors().accent || accent;
          doc.styles.tableHeader.fontSize = 8;
          doc.styles.title.fontSize = 13;
          doc.pageMargins = [20, 24, 20, 24];
          doc.footer = (page, pages) => ({ text: `Generated ${new Date().toLocaleString()}  ·  Page ${page} of ${pages}`, alignment: 'center', fontSize: 7, margin: [0, 6, 0, 0] });
        }
      },
      {
        extend: 'print', text: `${ICON.print}<span>Print</span>`, ...base,
        customize: win => {
          const s = win.document.createElement('style');
          s.textContent = `body{font-family:Segoe UI,Arial,sans-serif;font-size:11px} table{border-collapse:collapse;width:100%} th{background:${APEX.theme.colors().accent};color:#fff;padding:6px} td{padding:5px;border-bottom:1px solid #ddd} h1{font-size:18px}`;
          win.document.head.appendChild(s);
        }
      }
    ];
  }

  const layout = {
    topStart: ['pageLength', 'buttons'],
    topEnd: 'search',
    bottomStart: 'info',
    bottomEnd: 'paging'
  };

  const language = {
    search: '',
    searchPlaceholder: 'Search records…',
    lengthMenu: '_MENU_ per page',
    info: 'Showing _START_–_END_ of _TOTAL_',
    infoEmpty: 'No records',
    infoFiltered: '(filtered from _MAX_)',
    zeroRecords: 'No matching records found',
    emptyTable: 'No records available'
  };

  const ready = () => typeof window.DataTable === 'function';

  /* ---------- Container #6: Sub-status matrix ---------- */
  function renderMatrix(m) {
    if (!ready()) return;
    const table = $('#matrixTable');
    const empty = table.closest('.table-well').querySelector('.empty-state');

    if (matrixDT) { matrixDT.destroy(); matrixDT = null; }
    table.innerHTML = '';

    $('#matrixBadge').textContent = `${m.subs.length} sub-status${m.subs.length === 1 ? '' : 'es'} · ${m.rows.length} managers`;
    empty.hidden = !!m.rows.length;
    if (!m.rows.length) return;

    const max = Math.max(1, ...m.rows.flatMap(r => m.subs.map(s => r.counts[s] || 0)));
    table.innerHTML = `
      <thead><tr><th class="mx-name">Manager</th>${m.subs.map(s => `<th class="mx-num">${esc(s)}</th>`).join('')}<th class="mx-total">TOTAL</th></tr></thead>
      <tfoot><tr><th class="mx-name">GRAND TOTAL</th>${m.totals.map(t => `<th class="mx-num">${fmtNum(t)}</th>`).join('')}<th class="mx-total">${fmtNum(m.grand)}</th></tr></tfoot>`;

    const heat = (d, type) => {
      if (type !== 'display') return d;
      return d ? `<span class="heat" style="--a:${(0.12 + 0.6 * (d / max)).toFixed(2)}">${fmtNum(d)}</span>` : '<span class="zero">0</span>';
    };

    matrixDT = new DataTable(table, {
      data: m.rows.map(r => [r.manager, ...m.subs.map(s => r.counts[s] || 0), r.total]),
      columns: [
        { className: 'mx-name', render: (d, t) => (t === 'display' ? esc(d) : d) },
        ...m.subs.map(() => ({ className: 'mx-num', type: 'num', render: heat })),
        { className: 'mx-total', type: 'num', render: (d, t) => (t === 'display' ? fmtNum(d) : d) }
      ],
      order: [[m.subs.length + 1, 'desc']],
      pageLength: 10,
      lengthMenu: [5, 10, 25, 50, { label: 'All', value: -1 }],
      autoWidth: false,
      layout,
      language: { ...language, searchPlaceholder: 'Search managers…' },
      buttons: exportButtons('Sub-status Report', { footer: true, pageSize: m.subs.length > 8 ? 'A3' : 'A4' })
    });
  }

  /* ---------- Container #7: Detail records ---------- */
  const renderers = {
    date: (d, t) => {
      const isDate = d instanceof Date;
      if (t === 'sort' || t === 'type') return isDate ? d.getTime() : 0;
      return isDate ? fmtDate(d) : t === 'display' ? esc(d) : d || '';
    },
    number: (d, t) => (t === 'display' ? (d === '' || d === null || isNaN(d) ? '' : fmtNum(d)) : d === '' || isNaN(d) ? -1 : d),
    status: (d, t) => {
      if (t !== 'display') return d || blankLabel;
      const s = (d || '').toLowerCase();
      const cls = !s ? 'blank' : s === 'pending' || s === 'completed' ? s : 'other';
      return `<span class="badge badge-${cls}">${esc(d || 'Blank')}</span>`;
    },
    category: (d, t) => (t === 'display' ? (d ? `<span class="cat-chip cat-${slug(d)}">${esc(d)}</span>` : `<span class="zero">${blankLabel}</span>`) : d),
    long: (d, t) => (t === 'display' ? `<span class="trunc" title="${esc(d)}">${esc(d)}</span>` : d),
    text: (d, t) => (t === 'display' ? esc(d) : d)
  };

  function renderDetail(records, total) {
    if (!ready()) return;
    const summary = $('#detailSummary');
    if (summary) summary.textContent = records.length === total
      ? `Showing all ${fmtNum(total)} records`
      : `Showing ${fmtNum(records.length)} of ${fmtNum(total)} records (filtered)`;

    if (detailDT) {
      detailDT.clear().rows.add(records).draw(false);
      return;
    }

    detailDT = new DataTable('#detailTable', {
      data: records,
      columns: detailColumns.map(c => ({
        data: c.key,
        title: c.title,
        defaultContent: '',
        type: c.type === 'date' || c.type === 'number' ? 'num' : undefined,
        render: renderers[c.type] || renderers.text
      })),
      order: [[Math.max(0, detailColumns.findIndex(c => c.key === 'date')), 'desc']],
      pageLength: 10,
      lengthMenu: [10, 25, 50, 100, { label: 'All', value: -1 }],
      deferRender: true,
      autoWidth: false,
      layout,
      language: { ...language, emptyTable: 'No records match the selected filters' },
      buttons: exportButtons('Detail Records', { pageSize: 'A3' })
    });
  }

  return { renderMatrix, renderDetail, helpers: { exportButtons, layout, language } };
})();
