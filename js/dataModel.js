/* APEX Dashboard - aggregation / business logic */
window.APEX = window.APEX || {};

APEX.model = (() => {
  const { categories, statuses, blankLabel } = APEX.config;
  const { monthKey } = APEX.utils;

  const lc = v => (v || '').toLowerCase();
  const byCategory = (records, cat) => records.filter(r => lc(r.category) === cat);
  const label = v => v || blankLabel;

  function summarize(rows) {
    const s = { total: rows.length, completed: 0, pending: 0, blank: 0, other: 0 };
    rows.forEach(r => {
      const st = lc(r.status);
      if (!st) s.blank++;
      else if (st === statuses.completed) s.completed++;
      else if (st === statuses.pending) s.pending++;
      else s.other++;
    });
    return s;
  }

  function groupCount(rows, key) {
    const m = new Map();
    rows.forEach(r => m.set(label(r[key]), (m.get(label(r[key])) || 0) + 1));
    return [...m.entries()].map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
  }

  /* Monthly totals with gap-filling and month-over-month % change */
  function monthly(rows) {
    const counts = new Map();
    let undated = 0;
    rows.forEach(r => {
      if (!r.date) { undated++; return; }
      const k = monthKey(r.date);
      counts.set(k, (counts.get(k) || 0) + 1);
    });
    const keys = [...counts.keys()].sort();
    if (!keys.length) return { months: [], undated, current: null, previous: null };

    const months = [];
    let [y, m] = keys[0].split('-').map(Number);
    const last = keys[keys.length - 1];
    for (;;) {
      const k = `${y}-${String(m).padStart(2, '0')}`;
      months.push({ key: k, count: counts.get(k) || 0 });
      if (k === last) break;
      if (++m > 12) { m = 1; y++; }
    }
    months.forEach((mo, i) => {
      const prev = i ? months[i - 1].count : null;
      mo.change = prev ? ((mo.count - prev) / prev) * 100 : null;
    });

    return {
      months,
      undated,
      current: months[months.length - 1],
      previous: months.length > 1 ? months[months.length - 2] : null
    };
  }

  /* Manager x Sub-status matrix */
  function matrix(records) {
    const subs = [...new Set(records.map(r => label(r.subStatus)))].sort((a, b) => (a === blankLabel) - (b === blankLabel) || a.localeCompare(b));
    const byMgr = new Map();
    records.forEach(r => {
      const mgr = label(r.manager);
      const sub = label(r.subStatus);
      if (!byMgr.has(mgr)) byMgr.set(mgr, {});
      const c = byMgr.get(mgr);
      c[sub] = (c[sub] || 0) + 1;
    });
    const rows = [...byMgr.entries()].map(([manager, counts]) => ({
      manager,
      counts,
      total: Object.values(counts).reduce((a, b) => a + b, 0)
    }));
    const totals = subs.map(s => rows.reduce((a, r) => a + (r.counts[s] || 0), 0));
    return { subs, rows, totals, grand: totals.reduce((a, b) => a + b, 0) };
  }

  function filterValues(records, keys) {
    const out = {};
    keys.forEach(k => {
      out[k] = groupCount(records, k).sort((a, b) => (a.name === blankLabel) - (b.name === blankLabel) || a.name.localeCompare(b.name));
    });
    return out;
  }

  function build(records) {
    const pends = byCategory(records, categories.pends);
    const ageing = byCategory(records, categories.ageing);
    return {
      records,
      pends: summarize(pends),
      ageing: summarize(ageing),
      trend: monthly(pends),
      pendsByManager: groupCount(pends, 'manager'),
      ageingByManager: groupCount(ageing, 'manager'),
      matrix: matrix(records),
      filterValues: filterValues(records, APEX.config.filterKeys)
    };
  }

  return { build, label, filterValues };
})();
