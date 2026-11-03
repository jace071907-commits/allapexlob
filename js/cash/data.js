/* APEX Dashboard - CASH CONTROL tab: configuration, read-only parsing (cash.xlsb) and aggregation */
window.APEX = window.APEX || {};

APEX.cashConfig = Object.freeze({
  files: ['cash.xlsb', 'Cash.xlsb', 'cash.xlsx'],
  idbKey: 'cashHandle',
  pollInterval: 15000,
  blankLabel: '(Blank)',
  /* Newest "COR Data Run Date MM DD YYYY" sheet is used; falls back to any sheet with cor_seq_no in column S */
  sheetPrefix: 'COR Data Run Date',
  /* Fixed column positions (0-based): B, D, P, Q, S */
  cols: { manager: 1, ageCat: 3, domain: 15, associate: 16, cor: 18 },
  searchFields: ['cor', 'associate', 'domain', 'manager'],
  filters: [
    { key: 'manager', label: 'Manager Name' },
    { key: 'associate', label: 'Associate Name' },
    { key: 'domain', label: 'Associate Domain' },
    { key: 'ageCat', label: 'Age Category' },
    { key: 'cor', label: 'COR Number' }
  ],
  /* KPI age groups by the lower bound of the Age Category ("16-23 Days" -> 16) */
  ageKpis: [
    { key: 'a0', from: 0, to: 7 },
    { key: 'a1', from: 8, to: 15 },
    { key: 'a2', from: 16, to: 23 },
    { key: 'a3', from: 24, to: Infinity }
  ],
  topN: 10
});

APEX.cashData = (() => {
  const cfg = APEX.cashConfig;
  const { str } = APEX.utils;
  const B = cfg.blankLabel;

  const clean = v => str(v).replace(/&nbsp;/gi, ' ').replace(/\s+/g, ' ').trim();
  /* "AM37436 - Kim Jacobe" -> "Kim Jacobe" */
  const assocName = v => { const s = clean(v); const i = s.indexOf(' - '); return i > -1 ? s.slice(i + 3).trim() : s; };
  const ageLow = cat => { const m = str(cat).match(/\d+/); return m ? Number(m[0]) : null; };

  function sheetDate(name) {
    const m = str(name).match(/(\d{1,2})\D+(\d{1,2})\D+(\d{4})/);
    return m ? new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2])).getTime() : 0;
  }

  function pickSheet(names) {
    const runs = names.filter(n => str(n).toLowerCase().startsWith(cfg.sheetPrefix.toLowerCase()));
    return runs.sort((a, b) => sheetDate(b) - sheetDate(a))[0] || null;
  }

  function cellValue(ws, r, c) {
    const row = ws['!data'] ? ws['!data'][r] : null;
    const cell = row ? row[c] : ws[XLSX.utils.encode_cell({ r, c })];
    return cell && cell.v !== undefined ? cell.v : '';
  }

  function readSheet(ws) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const { manager, ageCat, domain, associate, cor } = cfg.cols;
    const header = Object.fromEntries(Object.entries(cfg.cols).map(([k, c]) => [k, clean(cellValue(ws, range.s.r, c))]));
    const records = [];
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const rec = {
        manager: clean(cellValue(ws, r, manager)),
        ageCat: clean(cellValue(ws, r, ageCat)),
        domain: clean(cellValue(ws, r, domain)),
        associate: assocName(cellValue(ws, r, associate)),
        cor: clean(cellValue(ws, r, cor))
      };
      if (!rec.cor && !rec.manager && !rec.associate) continue;
      rec.ageLow = ageLow(rec.ageCat);
      rec._search = cfg.searchFields.map(k => rec[k]).join(' \u0001 ').toLowerCase();
      records.push(rec);
    }
    return { records, header };
  }

  function parse(buffer) {
    if (typeof XLSX === 'undefined') throw new Error('SheetJS library failed to load.');
    const meta = XLSX.read(buffer, { type: 'array', bookSheets: true });
    let name = pickSheet(meta.SheetNames);
    let wb = name ? XLSX.read(buffer, { type: 'array', sheets: name, dense: true }) : null;

    if (!name) {
      /* fallback: first sheet whose column S header looks like the COR number */
      wb = XLSX.read(buffer, { type: 'array', dense: true });
      name = wb.SheetNames.find(n => /cor/i.test(str(cellValue(wb.Sheets[n], 0, cfg.cols.cor))) && /manager/i.test(str(cellValue(wb.Sheets[n], 0, cfg.cols.manager))));
      if (!name) throw new Error(`No "${cfg.sheetPrefix} …" sheet (or sheet with Manager in column B and COR number in column S) was found in the CASH workbook.`);
    }
    const { records, header } = readSheet(wb.Sheets[name]);
    return { records, sheet: name, header };
  }

  /* ---------- filtering ---------- */
  const val = (r, k) => r[k] || B;

  function applySelections(records, selections) {
    const active = Object.entries(selections).filter(([, s]) => s && s.size);
    return active.length ? records.filter(r => active.every(([k, s]) => s.has(val(r, k)))) : records;
  }

  function applySearch(records, query) {
    const terms = str(query).toLowerCase().split(/\s+/).filter(Boolean);
    return terms.length ? records.filter(r => terms.every(t => r._search.includes(t))) : records;
  }

  function countBy(records, key) {
    const m = new Map();
    records.forEach(r => m.set(val(r, key), (m.get(val(r, key)) || 0) + 1));
    return [...m.entries()].map(([name, count]) => ({ name, count }));
  }

  const byAge = (a, b) => (a.name === B) - (b.name === B) || (ageLow(a.name) ?? 1e9) - (ageLow(b.name) ?? 1e9) || a.name.localeCompare(b.name);
  const alpha = (a, b) => (a.name === B) - (b.name === B) || a.name.localeCompare(b.name, undefined, { numeric: true });
  const desc = (a, b) => b.count - a.count || a.name.localeCompare(b.name);

  const options = (records, key) => countBy(records, key).sort(key === 'ageCat' ? byAge : alpha);

  function build(records) {
    const distinct = k => new Set(records.map(r => r[k]).filter(Boolean)).size;
    const kpi = {
      total: records.length,
      unique: distinct('cor'),
      managers: distinct('manager'),
      associates: distinct('associate'),
      domains: distinct('domain')
    };
    cfg.ageKpis.forEach(a => { kpi[a.key] = records.filter(r => r.ageLow !== null && r.ageLow >= a.from && r.ageLow <= a.to).length; });
    kpi.noAge = records.filter(r => r.ageLow === null).length;

    const byManager = countBy(records, 'manager').sort(desc);
    const byAssociate = countBy(records, 'associate').sort(desc);
    return {
      kpi,
      byManager,
      byAssociate,
      byDomain: countBy(records, 'domain').sort(desc),
      byAge: countBy(records, 'ageCat').sort(byAge),
      topManagers: byManager.filter(x => x.name !== B).slice(0, cfg.topN),
      topAssociates: byAssociate.filter(x => x.name !== B).slice(0, cfg.topN)
    };
  }

  return { parse, applySelections, applySearch, options, build, ageLow };
})();
