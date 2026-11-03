/* APEX Dashboard - NON COB charts (Chart.js, reusing shared bar / tooltip helpers) */
window.APEX = window.APEX || {};

APEX.noncobCharts = (() => {
  const { $, fmtNum } = APEX.utils;
  const { topN } = APEX.noncobConfig;
  let pick = () => {};

  const PC = () => APEX.prepayCharts;
  const H = () => APEX.charts.helpers;
  const badge = (id, text) => { const el = $(`#${id}`); if (el) el.textContent = text; };
  const legend = c => ({ display: true, position: 'bottom', labels: { color: c.text, usePointStyle: true, boxWidth: 8, padding: 16, font: { weight: '600' } } });
  const valueScale = c => ({ beginAtZero: true, grace: '8%', grid: { color: c.grid }, ticks: { color: c.muted, precision: 0 }, border: { display: false } });
  const catScale = (c, color) => ({ grid: { display: false }, ticks: { color: color || c.muted, font: { weight: '600', size: 11.5 }, autoSkip: false }, border: { color: c.grid } });

  function doughnut(id, rows, c, filterKey) {
    PC().setEmpty(id, !rows.length);
    const palette = [c.accent, c.accent2, c.good, '#ffb547', c.bad, '#8d97ab', '#b48cff', '#4dd0e1'];
    const total = rows.reduce((a, r) => a + r.count, 0);
    PC().upsert(id, {
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
            textAlign: 'center',
            formatter: v => `${fmtNum(v)}\n${((v / total) * 100).toFixed(0)}%`
          }
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '58%',
        animation: { animateRotate: true, duration: 1200, easing: 'easeOutQuart' },
        onHover: (e, els) => { e.native.target.style.cursor = els.length && filterKey ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length && filterKey) pick(filterKey, rows[els[0].index].name); },
        plugins: {
          legend: { position: 'right', labels: { color: c.text, usePointStyle: true, boxWidth: 9, padding: 14, font: { weight: '600' } } },
          tooltip: {
            ...H().tooltip(c),
            callbacks: {
              label: item => ` ${item.label}: ${fmtNum(item.raw)} (${total ? ((item.raw / total) * 100).toFixed(1) : 0}%)`,
              footer: () => (filterKey ? 'Click to filter' : '')
            },
            footerColor: c.muted,
            footerFont: { size: 11, style: 'italic', weight: '400' }
          }
        }
      }
    });
  }

  const bar = (id, rows, c, key, opts = {}) => PC().bar(id, rows, c, { label: 'Records', ...opts, onPick: key ? v => pick(key, v) : null });

  /* Ageing: current age vs pend age per bucket */
  function ageingAge(rows, c) {
    const has = rows.some(r => r.curr || r.pend);
    PC().setEmpty('ncAgeChart', !has);
    PC().upsert('ncAgeChart', {
      type: 'bar',
      data: {
        labels: rows.map(r => r.name),
        datasets: [
          { label: 'Current Age (CURR_AGE)', data: rows.map(r => r.curr), backgroundColor: H().gradient(c, false, 0.55, 1), borderRadius: 8, datalabels: { display: ctx => ctx.dataset.data[ctx.dataIndex] > 0, anchor: 'end', align: 'end', color: c.textStrong, font: { weight: '700', size: 11 } } },
          { label: 'Pend Age (PEND_AGE_CURRENT)', data: rows.map(r => r.pend), backgroundColor: H().rgba(c.accent2Rgb, 0.6), borderRadius: 8, datalabels: { display: ctx => ctx.dataset.data[ctx.dataIndex] > 0, anchor: 'end', align: 'end', color: c.textStrong, font: { weight: '700', size: 11 } } }
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 22 } },
        interaction: { mode: 'index', intersect: false },
        animation: { duration: 1100, easing: 'easeOutQuart' },
        plugins: { legend: legend(c), tooltip: { ...H().tooltip(c), callbacks: { label: item => ` ${item.dataset.label}: ${fmtNum(item.raw)}` } } },
        scales: { x: catScale(c), y: valueScale(c) }
      }
    });
  }

  /* OPID: carry over vs non carry over per manager */
  function carryOver(rows, c) {
    PC().setEmpty('ncOCarryChart', !rows.length);
    const box = $('#ncOCarryChart')?.parentElement;
    if (box) box.style.height = `${Math.max(260, rows.length * 32 + 70)}px`;
    PC().upsert('ncOCarryChart', {
      type: 'bar',
      data: {
        labels: rows.map(r => r.name),
        datasets: [
          { label: 'Carry Over', data: rows.map(r => r.carry), backgroundColor: H().gradient(c, true, 0.55, 1), borderRadius: 6, stack: 's', datalabels: { display: ctx => ctx.dataset.data[ctx.dataIndex] > 0, color: '#fff', font: { weight: '700', size: 10.5 } } },
          { label: 'Not Carry Over', data: rows.map(r => r.non), backgroundColor: H().rgba(c.accent2Rgb, 0.45), borderRadius: 6, stack: 's', datalabels: { display: ctx => ctx.dataset.data[ctx.dataIndex] > 0, color: c.textStrong, font: { weight: '700', size: 10.5 } } }
        ]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        animation: { duration: 1100, easing: 'easeOutQuart' },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length) pick('manager', rows[els[0].index].name); },
        plugins: {
          legend: legend(c),
          tooltip: {
            ...H().tooltip(c),
            callbacks: {
              label: item => ` ${item.dataset.label}: ${fmtNum(item.raw)}`,
              footer: items => {
                const r = rows[items[0].dataIndex];
                const t = r.carry + r.non;
                return `Carry over rate: ${t ? ((r.carry / t) * 100).toFixed(1) : 0}% · click to filter`;
              }
            }
          }
        },
        scales: { x: { ...valueScale(c), stacked: true }, y: { ...catScale(c, c.text), stacked: true } }
      }
    });
  }

  function render(m) {
    if (!window.Chart || !APEX.prepayCharts) return;
    const c = APEX.theme.colors();
    const a = m.ageing, r = m.rstatus, o = m.opid;

    bar('ncAssocChart', a.byAssociate, c, 'associate');
    bar('ncTeamChart', a.byTeam, c, 'team', { alt: true });
    bar('ncMgrChart', a.byManager, c, 'manager');
    ageingAge(a.ageDist, c);
    doughnut('ncRecovChart', a.byRecovery, c, 'recoveryType');
    doughnut('ncStatusChart', a.byStatus, c, 'status');

    doughnut('ncRDistChart', r.byRStatus, c, 'rStatus');
    bar('ncRAssocChart', r.byAssociate, c, 'associate', { alt: true });
    bar('ncRTeamChart', r.byTeam, c, 'team');

    bar('ncOProcChart', o.byProcessor.slice(0, topN), c, 'processor');
    bar('ncOMgrChart', o.byManager, c, 'manager', { alt: true });
    bar('ncOMktChart', o.byMkt, c, null);
    carryOver(o.carryByManager, c);
    bar('ncOAgeChart', o.ageDist, c, null, { horizontal: false, alt: true });

    badge('ncAssocBadge', `${a.byAssociate.length} associates`);
    badge('ncRAssocBadge', `${r.byAssociate.length} associates`);
    badge('ncOProcBadge', `Top ${Math.min(topN, o.byProcessor.length)} of ${o.byProcessor.length}`);
    badge('ncOMgrBadge', `${o.byManager.length} managers`);
  }

  const onPick = fn => { pick = fn; };
  return { render, onPick };
})();
