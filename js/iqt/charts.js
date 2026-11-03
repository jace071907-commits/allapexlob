/* APEX Dashboard - IQT charts (click any bar / slice / point to filter the whole IQT tab) */
window.APEX = window.APEX || {};

APEX.iqtCharts = (() => {
  const { $, fmtNum } = APEX.utils;
  const { topN } = APEX.iqtConfig;
  let pick = () => {};

  const PC = () => APEX.prepayCharts;
  const H = () => APEX.charts.helpers;
  const badge = (id, text) => { const el = $(`#${id}`); if (el) el.textContent = text; };
  const bar = (id, rows, c, key, opts = {}) => PC().bar(id, rows, c, { label: 'Records', ...opts, onPick: v => pick(key, v) });

  function doughnut(id, rows, c, key) {
    PC().setEmpty(id, !rows.length);
    const palette = [c.accent, c.accent2, c.good, '#ffb547', c.bad, '#8d97ab'];
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
            display: ctx => total && ctx.dataset.data[ctx.dataIndex] / total >= 0.03,
            color: '#fff', font: { weight: '700', size: 11.5 }, textAlign: 'center',
            textShadowColor: 'rgba(0,0,0,.5)', textShadowBlur: 4,
            formatter: v => `${fmtNum(v)}\n${((v / total) * 100).toFixed(0)}%`
          }
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, cutout: '58%',
        animation: { animateRotate: true, duration: 1200, easing: 'easeOutQuart' },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length) pick(key, rows[els[0].index].name); },
        plugins: {
          legend: { position: 'right', labels: { color: c.text, usePointStyle: true, boxWidth: 9, padding: 14, font: { weight: '600' } } },
          tooltip: {
            ...H().tooltip(c),
            callbacks: { label: item => ` ${item.label}: ${fmtNum(item.raw)} (${total ? ((item.raw / total) * 100).toFixed(1) : 0}%)`, footer: () => 'Click to filter' },
            footerColor: c.muted, footerFont: { size: 11, style: 'italic', weight: '400' }
          }
        }
      }
    });
  }

  function pendedTrend(rows, c) {
    PC().setEmpty('iqPendedChart', !rows.length);
    PC().upsert('iqPendedChart', {
      type: 'line',
      data: {
        labels: rows.map(r => r.name.slice(0, 5)),
        datasets: [{
          label: 'Records',
          data: rows.map(r => r.count),
          borderColor: c.accent, borderWidth: 3, tension: 0.35, fill: true,
          backgroundColor: ctx => {
            const { chart } = ctx;
            if (!chart.chartArea) return H().rgba(c.accentRgb, 0.2);
            const g = chart.ctx.createLinearGradient(0, chart.chartArea.top, 0, chart.chartArea.bottom);
            g.addColorStop(0, H().rgba(c.accentRgb, 0.45));
            g.addColorStop(1, H().rgba(c.accentRgb, 0.02));
            return g;
          },
          pointBackgroundColor: c.accent, pointBorderColor: c.surface, pointBorderWidth: 2, pointRadius: 4, pointHoverRadius: 8,
          datalabels: { display: true, align: 'top', offset: 4, color: c.textStrong, font: { weight: '700', size: 10.5 } }
        }]
      },
      options: {
        responsive: true, maintainAspectRatio: false, layout: { padding: { top: 22 } },
        interaction: { mode: 'index', intersect: false },
        animation: { duration: 1200, easing: 'easeOutQuart' },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length) pick('pendedDay', rows[els[0].index].name); },
        plugins: {
          legend: { display: false },
          tooltip: { ...H().tooltip(c), callbacks: { title: items => rows[items[0].dataIndex].name, label: item => ` Records: ${fmtNum(item.raw)}`, footer: () => 'Click to filter' }, footerColor: c.muted, footerFont: { size: 11, style: 'italic', weight: '400' } }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: c.muted, font: { weight: '600' }, maxRotation: 0, autoSkip: true }, border: { color: c.grid } },
          y: { beginAtZero: true, grace: '12%', grid: { color: c.grid }, ticks: { color: c.muted, precision: 0 }, border: { display: false } }
        }
      }
    });
  }

  function render(m) {
    if (!window.Chart || !APEX.prepayCharts) return;
    const c = APEX.theme.colors();
    bar('iqSubChart', m.bySub, c, 'subStatus');
    bar('iqAssignedChart', m.byAssigned, c, 'assignedTo', { alt: true });
    bar('iqAgeChart', m.byAge, c, 'ageBucket', { horizontal: false });
    doughnut('iqStatusChart', m.byStatus, c, 'status');
    bar('iqOperatorChart', m.byOperator.slice(0, topN), c, 'operator', { alt: true });
    pendedTrend(m.daily, c);
    badge('iqSubBadge', `${m.bySub.length} sub-statuses`);
    badge('iqAssignedBadge', `${m.byAssigned.length} users`);
    badge('iqOperatorBadge', `Top ${Math.min(topN, m.byOperator.length)} of ${m.byOperator.length}`);
    badge('iqPendedBadge', m.daily.length ? `${m.daily[0].name} – ${m.daily[m.daily.length - 1].name}` : 'No dates');
  }

  const onPick = fn => { pick = fn; };
  return { render, onPick };
})();
