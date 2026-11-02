/* APEX Dashboard - NON COB workbook parsing (read-only) and aggregation */
window.APEX = window.APEX || {};

APEX.noncobData = (() => {
  const cfg = APEX.noncobConfig;
  const { str, toDate } = APEX.utils;
  const B = cfg.blankLabel;
  const SHEETS = Object.keys(cfg.sheets);

  const num = v => (v === '' || v === null || v === undefined || isNaN(Number(v)) ? null : Number(v));
  const flip = v => (v.includes(',') ? v.split(',').map(s => s.trim()).reverse().join(' ').replace(/\s+/g, ' ') : v);

  function bucketOf(value) {
    if (value === null) return '';
    const b = cfg.ageBuckets.find(([lo, hi]) => value >= lo && value <= hi);
    return b ? b[2] : '';
  }

  /* "16 - 30 DAYS" -> 16-30 Days; "120+ DAYS" -> 90+ Days; falls back to the numeric age */
  function ageBucket(cat, age) {
    const nums = (str(cat).match(/\d+/g) || []).map(Number);
    if (nums.length >= 2) return bucketOf(nums[1]);
    if (nums.length === 1) return bucketOf(str(cat).includes('+') ? nums[0] + 1 : nums[0]);
    return bucketOf(age);
  }

  function parseSheet(ws, sheetKey) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
    if (!rows.length) return { records: [], fields: [] };
    const lookup = new Map(rows[0].map((h, i) => [str(h).toLowerCase(), i]));
    const idx = {};
    Object.entries(cfg.columns[sheetKey]).forEach(([k, aliases]) => {
      const a = aliases.find(x => lookup.has(x.toLowerCase()));
      if (a) idx[k] = lookup.get(a.toLowerCase());
    });
    const flips = (cfg.flipNameFields[sheetKey] || []);

    const records = rows.slice(1).map(row => {
      const r = { _sheet: sheetKey };
      Object.entries(idx).forEach(([k, i]) => {
        const v = row[i];
        if (cfg.dateFields.includes(k)) r[k] = toDate(v);
        else if (cfg.numFields.includes(k)) r[k] = num(v);
        else {
          const s = str(v);
          r[k] = flips.includes(k) ? flip(s) : s;
        }
      });
      return r;
    }).filter(r => r.associate || r.claimId || r.claimNumber || r.errNum);

    if (sheetKey === 'ageing') records.forEach(r => { r.ageBucket = ageBucket(r.currAgeCat, r.currAge); r.pendBucket = ageBucket(r.pendAgeCat, r.pendAge); });
    if (sheetKey === 'opid') records.forEach(r => { r.carryOver = /carry\s*over/i.test(r.carryOver || '') ? 'Carry Over' : 'Not Carry Over'; });
    return { records, fields: Object.keys(idx).concat(sheetKey === 'opid' ? ['carryOver'] : []) };
  }

  function parse(buffer) {
    if (typeof XLSX === 'undefined') throw new Error('SheetJS library failed to load.');
    const meta = XLSX.read(buffer, { type: 'array', bookSheets: true });
    const resolve = name => meta.SheetNames.find(n => str(n).toLowerCase() === name.toLowerCase());
    const names = Object.fromEntries(SHEETS.map(k => [k, resolve(cfg.sheets[k])]));
    const found = Object.values(names).filter(Boolean);
    if (!found.length) throw new Error(`None of the sheets ${Object.values(cfg.sheets).map(s => `"${s}"`).join(', ')} were found in the NON COB workbook.`);

    const wb = XLSX.read(buffer, { type: 'array', sheets: found });
    const data = {};
    const fields = {};
    const missing = [];
    SHEETS.forEach(k => {
      if (!names[k]) { missing.push(cfg.sheets[k]); data[k] = []; fields[k] = []; return; }
      const res = parseSheet(wb.Sheets[names[k]], k);
      data[k] = res.records;
      fields[k] = res.fields;
    });
    return { data, fields, missing };
  }

  /* ---------- filtering (a filter only applies to sheets that contain the field) ---------- */
  const val = (r, k) => r[k] || B;

  function applyFilters(data, fields, selections) {
    const active = Object.entries(selections).filter(([, s]) => s && s.size);
    const out = {};
    SHEETS.forEach(sheet => {
      const rel = active.filter(([k]) => fields[sheet].includes(k));
      out[sheet] = rel.length ? data[sheet].filter(r => rel.every(([k, s]) => s.has(val(r, k)))) : data[sheet];
    });
    return out;
  }

  function filterOptions(data, fields) {
    const out = {};
    cfg.filters.forEach(({ key }) => {
      const m = new Map();
      SHEETS.forEach(sheet => {
        if (!fields[sheet].includes(key)) return;
        data[sheet].forEach(r => m.set(val(r, key), (m.get(val(r, key)) || 0) + 1));
      });
      out[key] = [...m.entries()].map(([name, count]) => ({ name, count }))
        .sort((a, b) => (a.name === B) - (b.name === B) || a.name.localeCompare(b.name));
    });
    return out;
  }

  /* ---------- aggregation ---------- */
  function countBy(rows, key) {
    const m = new Map();
    rows.forEach(r => m.set(val(r, key), (m.get(val(r, key)) || 0) + 1));
    return [...m.entries()].map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
  }

  function build(f) {
    const { ageing: A, rstatus: R, opid: O } = f;
    const buckets = cfg.ageBuckets.map(b => b[2]);
    const bucketCounts = Object.fromEntries(buckets.map(b => [b, A.filter(r => r.ageBucket === b).length]));
    const ages = O.map(r => r.ageDays).filter(v => v !== null);
    const carry = O.filter(r => r.carryOver === 'Carry Over').length;

    const carryByMgr = new Map();
    O.forEach(r => {
      const k = val(r, 'manager');
      if (!carryByMgr.has(k)) carryByMgr.set(k, { name: k, carry: 0, non: 0 });
      carryByMgr.get(k)[r.carryOver === 'Carry Over' ? 'carry' : 'non']++;
    });

    const opidAge = countBy(O, 'ageBucket');
    const order = [...cfg.opidAgeOrder, B];
    opidAge.sort((a, b) => (order.indexOf(a.name) + 1 || 99) - (order.indexOf(b.name) + 1 || 99));

    return {
      totals: { all: A.length + R.length + O.length, ageing: A.length, rstatus: R.length, opid: O.length },
      ageing: {
        buckets: bucketCounts,
        unbucketed: A.length - Object.values(bucketCounts).reduce((a, b) => a + b, 0),
        byAssociate: countBy(A, 'associate'),
        byTeam: countBy(A, 'team'),
        byManager: countBy(A, 'manager'),
        byRecovery: countBy(A, 'recoveryType'),
        byStatus: countBy(A, 'status'),
        ageDist: buckets.map(b => ({ name: b, curr: bucketCounts[b], pend: A.filter(r => r.pendBucket === b).length })),
        pend: Object.fromEntries(cfg.pendKpis.map(k => [k.key, A.filter(r => k.match.test(r.pendDesc || '')).length])),
        homePlan: A.filter(r => /waiting on home plan/i.test(r.pendDesc || '')).length
      },
      rstatus: {
        byRStatus: countBy(R, 'rStatus'),
        byAssociate: countBy(R, 'associate'),
        byTeam: countBy(R, 'team')
      },
      opid: {
        carry,
        nonCarry: O.length - carry,
        avgAge: ages.length ? ages.reduce((a, b) => a + b, 0) / ages.length : null,
        maxAge: ages.length ? Math.max(...ages) : null,
        agedCount: ages.length,
        byProcessor: countBy(O, 'processor'),
        byManager: countBy(O, 'manager'),
        byMkt: countBy(O, 'mktManager'),
        carryByManager: [...carryByMgr.values()].sort((a, b) => b.carry + b.non - (a.carry + a.non)),
        ageDist: opidAge
      }
    };
  }

  return { parse, applyFilters, filterOptions, build };
})();
