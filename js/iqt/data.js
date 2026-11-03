/* APEX Dashboard - IQT tab: configuration, read-only parsing and aggregation (source: IQT.xlsx › DUMP) */
window.APEX = window.APEX || {};

APEX.iqtConfig = Object.freeze({
  files: ['IQT.xlsx'],
  sheet: 'DUMP',
  idbKey: 'iqtHandle',
  pollInterval: 15000,
  blankLabel: '(Blank)',
  columns: {
    inquiry: ['Inquiry number', 'Inquiry Number'],
    totalAge: ['Total Age'],
    status: ['Status'],
    subStatus: ['Sub-status', 'Sub Status'],
    dcn: ['DCN'],
    pended: ['Pended Date'],
    assignedTo: ['Assigned To'],
    operator: ['Last Updated Operator ID'],
    lastNote: ['Last Note']
  },
  required: ['inquiry', 'status'],
  ageBuckets: [[0, 15, '0-15 Days'], [16, 30, '16-30 Days'], [31, 60, '31-60 Days'], [61, 90, '61-90 Days'], [91, Infinity, '91+ Days']],
  agedFrom: 61,
  searchFields: ['inquiry', 'dcn', 'assignedTo', 'operator', 'status', 'subStatus', 'lastNote'],
  globalFilters: [
    { key: 'status', label: 'Status' },
    { key: 'subStatus', label: 'Sub-status' },
    { key: 'assignedTo', label: 'Assigned To' },
    { key: 'operator', label: 'Last Updated Operator ID' },
    { key: 'ageBucket', label: 'Total Age Bucket' },
    { key: 'pendedDay', label: 'Pended Date' }
  ],
  tableFilters: [
    { key: 'status', label: 'Status' },
    { key: 'subStatus', label: 'Sub-status' },
    { key: 'assignedTo', label: 'Assigned To' },
    { key: 'ageBucket', label: 'Age Bucket' }
  ],
  topN: 15
});

APEX.iqtData = (() => {
  const cfg = APEX.iqtConfig;
  const { str, toDate, fmtDate } = APEX.utils;
  const B = cfg.blankLabel;

  const bucketOf = age => {
    if (age === null) return '';
    const b = cfg.ageBuckets.find(([lo, hi]) => age >= lo && age <= hi);
    return b ? b[2] : '';
  };

  function parseDateTime(v) {
    if (v instanceof Date) return isNaN(v) ? null : v;
    if (typeof v === 'number') return toDate(v);
    const s = str(v);
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d) ? toDate(s) : d;
  }

  function parse(buffer) {
    if (typeof XLSX === 'undefined') throw new Error('SheetJS library failed to load.');
    const meta = XLSX.read(buffer, { type: 'array', bookSheets: true });
    const name = meta.SheetNames.find(n => str(n).toLowerCase() === cfg.sheet.toLowerCase());
    if (!name) throw new Error(`Worksheet "${cfg.sheet}" was not found in the IQT workbook.`);
    const ws = XLSX.read(buffer, { type: 'array', sheets: name }).Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
    if (!rows.length) return { records: [] };

    const lookup = new Map(rows[0].map((h, i) => [str(h).toLowerCase(), i]));
    const idx = {};
    Object.entries(cfg.columns).forEach(([k, aliases]) => {
      const a = aliases.find(x => lookup.has(x.toLowerCase()));
      if (a !== undefined) idx[k] = lookup.get(a.toLowerCase());
    });
    const missing = cfg.required.filter(k => idx[k] === undefined);
    if (missing.length) throw new Error(`DUMP sheet is missing column(s): ${missing.map(k => cfg.columns[k][0]).join(', ')}`);

    const records = rows.slice(1).map(row => {
      const r = {};
      Object.entries(idx).forEach(([k, i]) => { r[k] = row[i]; });
      ['inquiry', 'status', 'subStatus', 'dcn', 'assignedTo', 'operator', 'lastNote'].forEach(k => { r[k] = str(r[k]); });
      const age = r.totalAge === '' || r.totalAge === null || r.totalAge === undefined || isNaN(Number(r.totalAge)) ? null : Number(r.totalAge);
      r.totalAge = age;
      r.ageBucket = bucketOf(age);
      r.pendedAt = parseDateTime(r.pended);
      r.pendedDate = r.pendedAt ? new Date(r.pendedAt.getFullYear(), r.pendedAt.getMonth(), r.pendedAt.getDate()) : null;
      r.pendedDay = r.pendedDate ? fmtDate(r.pendedDate) : '';
      r._search = cfg.searchFields.map(k => r[k] ?? '').join(' \u0001 ').toLowerCase();
      return r;
    }).filter(r => r.inquiry || r.dcn);

    return { records, sheet: name };
  }

  /* ---------- filtering ---------- */
  const val = (r, k) => r[k] || B;

  function applySelections(records, selections) {
    const active = Object.entries(selections).filter(([, s]) => s && s.size);
    return active.length ? records.filter(r => active.every(([k, s]) => s.has(val(r, k)))) : records;
  }

  /* All space-separated terms must match one of the search fields */
  function applySearch(records, query) {
    const terms = str(query).toLowerCase().split(/\s+/).filter(Boolean);
    return terms.length ? records.filter(r => terms.every(t => r._search.includes(t))) : records;
  }

  function applyDateRange(records, from, to) {
    if (!from && !to) return records;
    return records.filter(r => r.pendedDate && (!from || r.pendedDate >= from) && (!to || r.pendedDate <= to));
  }

  function countBy(records, key) {
    const m = new Map();
    records.forEach(r => m.set(val(r, key), (m.get(val(r, key)) || 0) + 1));
    return [...m.entries()].map(([name, count]) => ({ name, count }));
  }

  const bucketOrder = [...cfg.ageBuckets.map(b => b[2]), B];
  function sortOptions(list, key) {
    if (key === 'ageBucket') return list.sort((a, b) => bucketOrder.indexOf(a.name) - bucketOrder.indexOf(b.name));
    if (key === 'pendedDay') return list.sort((a, b) => (a.name === B) - (b.name === B) || new Date(a.name) - new Date(b.name));
    return list.sort((a, b) => (a.name === B) - (b.name === B) || a.name.localeCompare(b.name));
  }

  const options = (records, key) => sortOptions(countBy(records, key), key);
  const desc = list => list.sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));

  function build(records) {
    const ages = records.map(r => r.totalAge).filter(v => v !== null);
    const lc = v => v.toLowerCase();
    const daily = new Map();
    records.forEach(r => { if (r.pendedDate) daily.set(r.pendedDate.getTime(), (daily.get(r.pendedDate.getTime()) || 0) + 1); });

    return {
      kpi: {
        total: records.length,
        unique: new Set(records.map(r => r.inquiry)).size,
        investigation: records.filter(r => lc(r.status) === 'pending-investigation').length,
        review: records.filter(r => lc(r.status) === 'pending-review').length,
        avgAge: ages.length ? ages.reduce((a, b) => a + b, 0) / ages.length : null,
        maxAge: ages.length ? Math.max(...ages) : null,
        aged: records.filter(r => r.totalAge !== null && r.totalAge >= cfg.agedFrom).length,
        assignees: new Set(records.map(r => r.assignedTo).filter(Boolean)).size,
        notPended: records.filter(r => !r.pendedDate).length
      },
      bySub: desc(countBy(records, 'subStatus')),
      byAssigned: desc(countBy(records, 'assignedTo')),
      byOperator: desc(countBy(records, 'operator')),
      byStatus: desc(countBy(records, 'status')),
      byAge: cfg.ageBuckets.map(b => ({ name: b[2], count: records.filter(r => r.ageBucket === b[2]).length })),
      daily: [...daily.entries()].sort((a, b) => a[0] - b[0]).map(([t, count]) => ({ date: new Date(t), name: fmtDate(new Date(t)), count }))
    };
  }

  return { parse, applySelections, applySearch, applyDateRange, options, build };
})();
