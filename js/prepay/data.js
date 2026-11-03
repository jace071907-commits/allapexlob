/* APEX Dashboard - PREPAY workbook parsing (read-only) and aggregation model */
window.APEX = window.APEX || {};

APEX.prepayData = (() => {
  const cfg = APEX.prepayConfig;
  const { str, toDate, fmtDate } = APEX.utils;
  const B = cfg.blankLabel;
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const isExcluded = name => cfg.excludeSheets.some(x => x.toLowerCase() === str(name).toLowerCase());
  const num = v => (v === '' || v === null || v === undefined || isNaN(Number(v)) ? null : Number(v));
  const isFinalized = u => /finali[sz]ed/i.test(u) && !/waiting/i.test(u);

  function columnIndex(headerRow) {
    const lookup = new Map(headerRow.map((h, i) => [str(h).toLowerCase(), i]));
    const idx = {};
    Object.entries(cfg.columns).forEach(([key, aliases]) => {
      const a = aliases.find(x => lookup.has(x.toLowerCase()));
      if (a) idx[key] = lookup.get(a.toLowerCase());
    });
    return idx;
  }

  /* ---------- Parse every October snapshot sheet into one dataset ---------- */
  function parse(buffer) {
    if (typeof XLSX === 'undefined') throw new Error('SheetJS library failed to load.');
    const meta = XLSX.read(buffer, { type: 'array', bookSheets: true });
    const names = meta.SheetNames.filter(n => !isExcluded(n) && cfg.sheetPattern.test(n));
    if (!names.length) throw new Error('No October snapshot sheets (e.g. 1001, 1002 …) were found in the PREPAY workbook.');
    const jiraName = cfg.jiraSheets.map(j => meta.SheetNames.find(n => str(n).toLowerCase() === j.toLowerCase())).find(Boolean);

    const wb = XLSX.read(buffer, { type: 'array', sheets: jiraName ? [...names, jiraName] : names });
    const raw = [];
    const skipped = [];
    const sheetDays = {};

    names.forEach(name => {
      const ws = wb.Sheets[name];
      const rows = ws ? XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false }) : [];
      const hi = rows.findIndex(r => r.some(c => str(c).toLowerCase() === 'dcn'));
      if (hi < 0) { skipped.push(name); return; }
      const idx = columnIndex(rows[hi]);
      if (cfg.requiredColumns.some(k => idx[k] === undefined)) { skipped.push(name); return; }
      const day = Number(cfg.sheetPattern.exec(name)[1]);
      sheetDays[name] = day;

      for (let i = hi + 1; i < rows.length; i++) {
        const row = rows[i];
        const rec = { _sheet: name, _day: day };
        Object.entries(idx).forEach(([k, c]) => { rec[k] = row[c]; });
        Object.keys(cfg.columns).forEach(k => {
          const v = rec[k];
          if (cfg.numericColumns.includes(k)) rec[k] = num(v);
          else if (cfg.dateColumns.includes(k)) rec[k] = toDate(v);
          else {
            const s = str(v);
            rec[k] = cfg.zeroAsBlank.includes(k) && s === '0' ? '' : s;
          }
        });
        if (!rec.dcn || rec.dcn === '0') continue;
        raw.push(rec);
      }
    });

    /* Sheet names carry month/day only – infer the year from the data */
    const years = {};
    raw.forEach(r => { if (r.lastResearch) years[r.lastResearch.getFullYear()] = (years[r.lastResearch.getFullYear()] || 0) + 1; });
    const year = Number(Object.entries(years).sort((a, b) => b[1] - a[1])[0]?.[0]) || new Date().getFullYear();

    raw.forEach(r => {
      r.snapshotDate = new Date(year, cfg.snapshotMonth - 1, r._day);
      r.snapshot = fmtDate(r.snapshotDate);
      r.finalized = isFinalized(r.updates);
      r.status = r.finalized ? 'Finalized' : 'Open';
      r.high = r.urgency !== null && r.urgency >= cfg.highUrgency;
      r.wgsCodes = [...new Set(r.wgs.split(/[,;\s]+/).map(s => s.trim().toUpperCase()).filter(s => s && s !== '0'))];
      r.dmBucket = bucket(r.dmAge, cfg.dmBuckets);
      r.touchBucket = bucket(r.lastTouched, cfg.touchBuckets);
    });

    const sheets = Object.entries(sheetDays)
      .map(([name, day]) => ({ name, day, date: new Date(year, cfg.snapshotMonth - 1, day) }))
      .sort((a, b) => a.day - b.day);

    const jira = jiraName && APEX.prepayJira ? APEX.prepayJira.parse(wb.Sheets[jiraName]) : null;

    return { records: raw, sheets, skipped, jira, jiraSheet: jiraName || null };
  }

  function bucket(v, defs) {
    if (v === null) return 'Not recorded';
    const b = defs.find(([lo, hi]) => v >= lo && v <= hi);
    return b ? b[2] : 'Not recorded';
  }

  /* ---------- Filtering ---------- */
  const filterValue = (r, key) => {
    if (key === 'wgs') return r.wgsCodes.length ? r.wgsCodes : [B];
    if (key === 'urgency') return [r.urgency === null ? B : String(r.urgency)];
    return [r[key] || B];
  };

  function applyFilters(records, selections) {
    const active = Object.entries(selections).filter(([, s]) => s && s.size);
    if (!active.length) return records;
    return records.filter(r => active.every(([k, s]) => filterValue(r, k).some(v => s.has(v))));
  }

  function countBy(records, fn) {
    const m = new Map();
    records.forEach(r => {
      const keys = fn(r);
      (Array.isArray(keys) ? keys : [keys]).forEach(k => m.set(k, (m.get(k) || 0) + 1));
    });
    return [...m.entries()].map(([name, count]) => ({ name, count }));
  }

  const alpha = (a, b) => (a.name === B) - (b.name === B) || a.name.localeCompare(b.name, undefined, { numeric: true });
  const desc = (a, b) => b.count - a.count || alpha(a, b);

  function filterOptions(records) {
    const out = {};
    cfg.filters.forEach(({ key }) => {
      const list = countBy(records, r => filterValue(r, key));
      if (key === 'urgency') list.sort((a, b) => (a.name === B) - (b.name === B) || Number(b.name) - Number(a.name));
      else if (key === 'snapshot') list.sort((a, b) => new Date(a.name) - new Date(b.name));
      else list.sort(alpha);
      out[key] = list;
    });
    return out;
  }

  /* ---------- Group summaries (used by reports) ---------- */
  function summarize(name, rows) {
    const dm = rows.map(r => r.dmAge).filter(v => v !== null);
    const finalized = rows.filter(r => r.finalized).length;
    return {
      name,
      total: rows.length,
      finalized,
      open: rows.length - finalized,
      valid: rows.filter(r => r.verified.toLowerCase() === 'valid').length,
      invalid: rows.filter(r => r.verified.toLowerCase() === 'invalid').length,
      high: rows.filter(r => r.high).length,
      avgDm: dm.length ? dm.reduce((a, b) => a + b, 0) / dm.length : null,
      maxDm: dm.length ? Math.max(...dm) : null,
      finalizedPct: rows.length ? (finalized / rows.length) * 100 : 0,
      rows
    };
  }

  function groupSummary(records, fn) {
    const m = new Map();
    records.forEach(r => {
      const keys = fn(r);
      (Array.isArray(keys) ? keys : [keys]).forEach(k => {
        if (!m.has(k)) m.set(k, []);
        m.get(k).push(r);
      });
    });
    return [...m.entries()].map(([k, rows]) => summarize(k, rows)).sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));
  }

  /* ---------- Dashboard model ---------- */
  function build(records) {
    const s = summarize('All', records);
    const snapMap = new Map();
    records.forEach(r => {
      if (!snapMap.has(r.snapshot)) snapMap.set(r.snapshot, { label: r.snapshot, date: r.snapshotDate, total: 0, finalized: 0 });
      const x = snapMap.get(r.snapshot);
      x.total++;
      if (r.finalized) x.finalized++;
    });
    const snapshots = [...snapMap.values()].sort((a, b) => a.date - b.date).map(x => ({ ...x, open: x.total - x.finalized, short: `${MONTHS[x.date.getMonth()]} ${String(x.date.getDate()).padStart(2, '0')}` }));

    const dayMap = new Map();
    records.forEach(r => {
      if (!r.lastResearch) return;
      const k = r.lastResearch.getTime();
      dayMap.set(k, (dayMap.get(k) || 0) + 1);
    });
    const daily = [...dayMap.entries()].sort((a, b) => a[0] - b[0]).map(([t, count]) => ({ date: new Date(t), count }));

    const bucketOrder = defs => [...defs.map(d => d[2]), 'Not recorded'];
    const ordered = (list, order) => order.map(name => ({ name, count: (list.find(x => x.name === name) || {}).count || 0 })).filter(x => x.count || x.name !== 'Not recorded');

    return {
      records,
      kpi: {
        total: s.total,
        unique: new Set(records.map(r => r.dcn)).size,
        finalized: s.finalized,
        valid: s.valid,
        open: s.open,
        avgDm: s.avgDm,
        maxDm: s.maxDm,
        high: s.high,
        latest: snapshots.length ? snapshots[snapshots.length - 1] : null
      },
      snapshots,
      daily,
      byAssigned: countBy(records, r => r.assignedTo || B).sort(desc),
      bySub: countBy(records, r => r.subStatus || B).sort(desc),
      bySkill: countBy(records, r => r.skill || B).sort(desc),
      byWgs: countBy(records, r => filterValue(r, 'wgs')).sort(desc),
      byUrgency: countBy(records, r => (r.urgency === null ? B : String(r.urgency))).sort((a, b) => (a.name === B) - (b.name === B) || Number(a.name) - Number(b.name)),
      byDm: ordered(countBy(records, r => r.dmBucket), bucketOrder(cfg.dmBuckets))
    };
  }

  return { parse, build, applyFilters, filterOptions, groupSummary, summarize };
})();
