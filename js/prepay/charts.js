/* APEX Dashboard - PREPAY charts (Chart.js) */
window.APEX = window.APEX || {};

APEX.prepayCharts = (() => {
  const { $, fmtNum, fmtDate } = APEX.utils;
  const { topN } = APEX.prepayConfig;
  const charts = {};
  let drill = () => {};

  const H = () => APEX.charts.helpers;
  const anim = { duration: 1100, easing: 'easeOutQuart', delay: ctx => (ctx.type === 'data' ? Math.min(ctx.dataIndex * 50, 800) : 0) };

  function upsert(id, config) {
    const canvas = $(`#${id}`);
    if (!canvas || !window.Chart) return;
    if (charts[id]) { charts[id].data = config.data; charts[id].options = config.options; charts[id].update(); }
    else charts[id] = new Chart(canvas, config);
  }

  function setEmpty(id, empty) {
    const es = $(`#${id}`)?.closest('.chart-well')?.querySelector('.empty-state');
    if (es) es.hidden = !empty;
  }

  const setBadge = (id, text) => { const el = $(`#${id}`); if (el) el.textContent = text; };

  const scales = (c, horizontal) => {
    const val = { beginAtZero: true, grace: '8%', grid: { color: c.grid }, ticks: { color: c.muted, precision: 0 }, border: { display: false } };
    const cat = { grid: { display: false }, ticks: { color: c.text, font: { weight: '600', size: 11.5 }, autoSkip: false }, border: { color: c.grid } };
    return horizontal ? { x: val, y: cat } : { x: { ...cat, ticks: { ...cat.ticks, color: c.muted, autoSkip: true, maxRotation: 0 } }, y: val };
  };

  /* Generic categorical bar chart */
  function barChart(id, rows, c, { horizontal = true, alt = false, label = 'Records', drillKey = null, total = null, onPick = null } = {}) {
    const pick = onPick || (drillKey ? v => drill(drillKey, v) : null);
    setEmpty(id, !rows.length);
    const box = $(`#${id}`)?.parentElement;
    if (box && horizontal) box.style.height = `${Math.max(260, rows.length * 30 + 50)}px`;
    const sum = total ?? rows.reduce((a, r) => a + r.count, 0);
    const cc = alt ? { ...c, accent: c.accent2, accentRgb: c.accent2Rgb, accent2Rgb: c.accentRgb } : c;

    upsert(id, {
      type: 'bar',
      data: {
        labels: rows.map(r => r.name),
        datasets: [{
          label,
          data: rows.map(r => r.count),
          backgroundColor: H().gradient(cc, horizontal, 0.45, 1),
          hoverBackgroundColor: cc.accent,
          borderRadius: 8,
          borderSkipped: false,
          barPercentage: 0.78,
          categoryPercentage: 0.88,
          datalabels: {
            display: true,
            anchor: 'end',
            align: 'end',
            offset: 4,
            color: c.textStrong,
            font: { weight: '700', size: 11 },
            formatter: v => fmtNum(v)
          }
        }]
      },
      options: {
        indexAxis: horizontal ? 'y' : 'x',
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: horizontal ? { right: 50 } : { top: 22 } },
        animation: anim,
        onHover: (e, els) => { e.native.target.style.cursor = els.length && pick ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length && pick) pick(rows[els[0].index].name); },
        plugins: {
          legend: { display: false },
          tooltip: {
            ...H().tooltip(c),
            callbacks: {
              label: item => ` ${label}: ${fmtNum(item.raw)}`,
              afterLabel: item => (sum ? ` Share: ${((item.raw / sum) * 100).toFixed(1)}%` : ''),
              footer: () => (pick ? (onPick ? 'Click to filter' : 'Click to filter the PREPAY dashboard') : '')
            },
            footerColor: c.muted,
            footerFont: { size: 11, style: 'italic', weight: '400' }
          }
        },
        scales: scales(c, horizontal)
      }
    });
  }

  function snapshotChart(snaps, c) {
    setEmpty('ppSnapshotChart', !snaps.length);
    upsert('ppSnapshotChart', {
      type: 'bar',
      data: {
        labels: snaps.map(s => s.short),
        datasets: [
          {
            type: 'line', label: 'Total', data: snaps.map(s => s.total), stack: 'total',
            borderColor: c.accent2, backgroundColor: c.accent2, pointBackgroundColor: c.accent2, pointBorderColor: c.surface,
            pointBorderWidth: 2, pointRadius: 4, pointHoverRadius: 7, borderWidth: 3, tension: 0.35, order: 0,
            datalabels: { display: true, align: 'top', anchor: 'end', offset: 4, color: c.textStrong, font: { weight: '700', size: 10.5 }, formatter: v => fmtNum(v) }
          },
          { type: 'bar', label: 'Finalized', data: snaps.map(s => s.finalized), stack: 'status', backgroundColor: H().rgba(c.accent2Rgb, 0.55), borderRadius: 6, order: 1 },
          { type: 'bar', label: 'Open', data: snaps.map(s => s.open), stack: 'status', backgroundColor: H().gradient(c, false, 0.6, 1), borderRadius: 6, order: 1 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        layout: { padding: { top: 24 } },
        animation: anim,
        plugins: {
          legend: { display: true, position: 'bottom', labels: { color: c.text, usePointStyle: true, boxWidth: 8, padding: 16 } },
          tooltip: {
            ...H().tooltip(c),
            callbacks: {
              title: items => `Snapshot ${fmtDate(snaps[items[0].dataIndex].date)}`,
              label: item => ` ${item.dataset.label}: ${fmtNum(item.raw)}`,
              footer: items => {
                const i = items[0].dataIndex;
                if (!i) return 'Baseline snapshot';
                const prev = snaps[i - 1].total;
                const ch = prev ? ((snaps[i].total - prev) / prev) * 100 : null;
                return ch === null ? '' : `vs previous snapshot: ${ch >= 0 ? '+' : ''}${ch.toFixed(1)}%`;
              }
            }
          }
        },
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { color: c.muted, font: { weight: '600' }, maxRotation: 0, autoSkip: true }, border: { color: c.grid } },
          y: { stacked: true, beginAtZero: true, grace: '10%', grid: { color: c.grid }, ticks: { color: c.muted, precision: 0 }, border: { display: false } }
        }
      }
    });
    setBadge('ppSnapshotBadge', `${snaps.length} snapshot${snaps.length === 1 ? '' : 's'}`);
  }

  function dailyChart(daily, c) {
    setEmpty('ppDailyChart', !daily.length);
    upsert('ppDailyChart', {
      type: 'line',
      data: {
        labels: daily.map(d => fmtDate(d.date).slice(0, 5)),
        datasets: [{
          label: 'Records',
          data: daily.map(d => d.count),
          borderColor: c.accent,
          borderWidth: 3,
          tension: 0.35,
          fill: true,
          backgroundColor: ctx => {
            const { chart } = ctx;
            if (!chart.chartArea) return H().rgba(c.accentRgb, 0.2);
            const g = chart.ctx.createLinearGradient(0, chart.chartArea.top, 0, chart.chartArea.bottom);
            g.addColorStop(0, H().rgba(c.accentRgb, 0.45));
            g.addColorStop(1, H().rgba(c.accentRgb, 0.02));
            return g;
          },
          pointBackgroundColor: c.accent,
          pointBorderColor: c.surface,
          pointBorderWidth: 2,
          pointRadius: 3,
          pointHoverRadius: 7
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        animation: { duration: 1200, easing: 'easeOutQuart' },
        plugins: {
          legend: { display: false },
          tooltip: { ...H().tooltip(c), callbacks: { title: items => fmtDate(daily[items[0].dataIndex].date), label: item => ` Records: ${fmtNum(item.raw)}` } }
        },
        scales: scales(c, false)
      }
    });
    setBadge('ppDailyBadge', daily.length ? `${fmtDate(daily[0].date)} – ${fmtDate(daily[daily.length - 1].date)}` : 'No dates');
  }

  function render(m) {
    if (!window.Chart) return;
    const c = APEX.theme.colors();
    snapshotChart(m.snapshots, c);
    dailyChart(m.daily, c);
    barChart('ppAssignedChart', m.byAssigned, c, { drillKey: 'assignedTo', label: 'Records', total: m.kpi.total });
    barChart('ppSubChart', m.bySub, c, { alt: true, drillKey: 'subStatus', total: m.kpi.total });
    barChart('ppSkillChart', m.bySkill, c, { drillKey: 'skill', total: m.kpi.total });
    barChart('ppWgsChart', m.byWgs.slice(0, topN), c, { alt: true, drillKey: 'wgs', total: m.kpi.total });
    barChart('ppUrgencyChart', m.byUrgency, c, { horizontal: false, drillKey: 'urgency', total: m.kpi.total });
    barChart('ppDmChart', m.byDm, c, { horizontal: false, alt: true, total: m.kpi.total });

    setBadge('ppAssignedBadge', `${m.byAssigned.length} associates`);
    setBadge('ppSubBadge', `${m.bySub.length} sub-statuses`);
    setBadge('ppSkillBadge', `${m.bySkill.length} skills`);
    setBadge('ppWgsBadge', `Top ${Math.min(topN, m.byWgs.length)} of ${m.byWgs.length} codes`);
  }

  const onDrill = fn => { drill = fn; };
  const resize = () => Object.values(charts).forEach(ch => ch.resize());

  return { render, onDrill, resize, bar: barChart, upsert, setEmpty };
})();
