/* APEX Dashboard - Chart.js visuals (Containers #2, #4, #5) */
window.APEX = window.APEX || {};

APEX.charts = (() => {
  const { $, esc, fmtNum, fmtPct, monthLabel, animateCount } = APEX.utils;
  const charts = {};
  let onBarClick = () => {};

  if (window.Chart) {
    if (window.ChartDataLabels) Chart.register(ChartDataLabels);
    Chart.defaults.font.family = '"Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif';
    Chart.defaults.plugins.datalabels = Object.assign(Chart.defaults.plugins.datalabels || {}, { display: false });
  }

  const rgba = (rgb, a) => `rgba(${rgb}, ${a})`;

  function tooltip(c) {
    return {
      backgroundColor: c.mode === 'light' ? 'rgba(255,255,255,0.97)' : 'rgba(18,21,28,0.96)',
      titleColor: c.textStrong,
      bodyColor: c.text,
      borderColor: rgba(c.accentRgb, 0.6),
      borderWidth: 1,
      padding: 12,
      cornerRadius: 10,
      titleFont: { weight: '700', size: 13 },
      bodyFont: { size: 12.5 },
      displayColors: true,
      boxPadding: 4,
      caretSize: 7
    };
  }

  /* Gradient fill that follows the chart area */
  const gradient = (c, horizontal, alphaFrom = 0.55, alphaTo = 1) => ctx => {
    const { chart } = ctx;
    const area = chart.chartArea;
    if (!area) return c.accent;
    const g = horizontal
      ? chart.ctx.createLinearGradient(area.left, 0, area.right, 0)
      : chart.ctx.createLinearGradient(0, area.bottom, 0, area.top);
    g.addColorStop(0, rgba(c.accent2Rgb, alphaFrom));
    g.addColorStop(1, rgba(c.accentRgb, alphaTo));
    return g;
  };

  function upsert(id, config) {
    const canvas = $(`#${id}`);
    if (!canvas || !window.Chart) return;
    if (charts[id]) {
      charts[id].data = config.data;
      charts[id].options = config.options;
      charts[id].update();
    } else {
      charts[id] = new Chart(canvas, config);
    }
  }

  function setEmpty(canvasId, empty) {
    const well = $(`#${canvasId}`)?.closest('.chart-well');
    const es = well && well.querySelector('.empty-state');
    if (es) es.hidden = !empty;
  }

  /* ---------- Container #2: Monthly trend ---------- */
  function renderTrend(trend, c) {
    const { months, current, previous, undated } = trend;
    setEmpty('trendChart', !months.length);

    $('#trendCurLabel').textContent = current ? `Current · ${monthLabel(current.key)}` : 'Current Month';
    $('#trendPrevLabel').textContent = previous ? `Previous · ${monthLabel(previous.key)}` : 'Previous Month';
    animateCount($('#trendCur'), current ? current.count : 0);
    animateCount($('#trendPrev'), previous ? previous.count : 0);

    const mom = current && previous ? current.change : null;
    const dir = mom === null ? '' : mom > 0 ? 'up' : mom < 0 ? 'down' : '';
    const arrow = dir === 'up' ? '▲' : dir === 'down' ? '▼' : '■';
    const momEl = $('#trendMom');
    momEl.textContent = mom === null ? 'N/A' : `${arrow} ${fmtPct(mom)}`;
    momEl.className = `mini-value ${dir}`;
    const badge = $('#momBadge');
    badge.textContent = mom === null ? 'MoM N/A' : `${arrow} ${fmtPct(mom)} MoM`;
    badge.className = `mom-badge ${dir}`;

    $('#momChips').innerHTML = months.map((m, i) => {
      const d = m.change === null ? 'flat' : m.change > 0 ? 'up' : m.change < 0 ? 'down' : 'flat';
      const txt = m.change === null ? (i === 0 ? 'baseline' : 'N/A') : `${d === 'up' ? '▲ Increase' : d === 'down' ? '▼ Decrease' : '■ No change'} ${fmtPct(m.change)}`;
      return `<span class="mom-chip ${d}" style="animation-delay:${i * 60}ms"><b>${esc(monthLabel(m.key))}</b> ${fmtNum(m.count)} <em>${esc(txt)}</em></span>`;
    }).join('');

    $('#trendNote').textContent = undated ? `${fmtNum(undated)} Pends record(s) without a Date are excluded from the trend.` : '';

    const labels = months.map(m => monthLabel(m.key));
    const data = months.map(m => m.count);

    upsert('trendChart', {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            type: 'line',
            label: 'Trend',
            data,
            borderColor: c.accent2,
            backgroundColor: rgba(c.accent2Rgb, 0.15),
            pointBackgroundColor: c.accent2,
            pointBorderColor: c.surface,
            pointBorderWidth: 2,
            pointRadius: 5,
            pointHoverRadius: 8,
            borderWidth: 3,
            tension: 0.4,
            fill: false,
            order: 0
          },
          {
            type: 'bar',
            label: 'Pends',
            data,
            backgroundColor: gradient(c, false),
            hoverBackgroundColor: c.accent,
            borderRadius: 10,
            borderSkipped: false,
            maxBarThickness: 64,
            order: 1,
            datalabels: {
              display: true,
              anchor: 'end',
              align: 'end',
              offset: 2,
              color: c.textStrong,
              font: { weight: '700', size: 12 },
              formatter: v => fmtNum(v)
            }
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        layout: { padding: { top: 24 } },
        animation: { duration: 1200, easing: 'easeOutQuart', delay: ctx => (ctx.type === 'data' ? ctx.dataIndex * 90 : 0) },
        plugins: {
          legend: { display: false },
          tooltip: {
            ...tooltip(c),
            filter: item => item.dataset.type === 'bar',
            callbacks: {
              label: item => ` Pends: ${fmtNum(item.raw)}`,
              afterLabel: item => {
                const ch = months[item.dataIndex].change;
                if (ch === null) return item.dataIndex === 0 ? ' Baseline month' : ' Change: N/A';
                return ` ${ch >= 0 ? 'Increase' : 'Decrease'} vs prev: ${fmtPct(ch)}`;
              }
            }
          }
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: c.muted, font: { weight: '600' } }, border: { color: c.grid } },
          y: { beginAtZero: true, grace: '10%', grid: { color: c.grid }, ticks: { color: c.muted, precision: 0 }, border: { display: false } }
        }
      }
    });
  }

  /* ---------- Containers #4 / #5: Manager bars ---------- */
  function renderManager(id, rows, c, category, alt) {
    setEmpty(id, !rows.length);
    const box = $(`#${id}`)?.parentElement;
    if (box) box.style.height = `${Math.max(260, rows.length * 46 + 40)}px`;

    const total = rows.reduce((a, r) => a + r.count, 0);
    const cc = alt ? { ...c, accent: c.accent2, accentRgb: c.accent2Rgb, accent2Rgb: c.accentRgb } : c;

    upsert(id, {
      type: 'bar',
      data: {
        labels: rows.map(r => r.name),
        datasets: [{
          label: `${category} Count`,
          data: rows.map(r => r.count),
          backgroundColor: gradient(cc, true, 0.45, 1),
          hoverBackgroundColor: cc.accent,
          borderColor: rgba(cc.accentRgb, 0.9),
          borderWidth: { right: 2 },
          borderRadius: 10,
          borderSkipped: false,
          barPercentage: 0.75,
          categoryPercentage: 0.85,
          datalabels: {
            display: true,
            anchor: 'end',
            align: 'end',
            offset: 6,
            color: c.textStrong,
            font: { weight: '700', size: 12 },
            formatter: v => `${fmtNum(v)}  (${total ? ((v / total) * 100).toFixed(1) : 0}%)`
          }
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { right: 90 } },
        animation: { duration: 1100, easing: 'easeOutQuart', delay: ctx => (ctx.type === 'data' ? ctx.dataIndex * 80 : 0) },
        onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        onClick: (e, els) => { if (els.length) onBarClick({ category, manager: rows[els[0].index].name }); },
        plugins: {
          legend: { display: false },
          tooltip: {
            ...tooltip(c),
            callbacks: {
              title: items => items[0].label,
              label: item => ` ${category} Count: ${fmtNum(item.raw)}`,
              afterLabel: item => ` Share: ${total ? ((item.raw / total) * 100).toFixed(1) : 0}% of ${fmtNum(total)}`,
              footer: () => 'Click to view detail records'
            },
            footerColor: c.muted,
            footerFont: { size: 11, style: 'italic', weight: '400' }
          }
        },
        scales: {
          x: { beginAtZero: true, grid: { color: c.grid }, ticks: { color: c.muted, precision: 0 }, border: { display: false } },
          y: { grid: { display: false }, ticks: { color: c.text, font: { weight: '600', size: 12.5 } }, border: { color: c.grid } }
        }
      }
    });

    const badge = $(`#${id === 'pendsMgrChart' ? 'pendsMgrCount' : 'ageingMgrCount'}`);
    if (badge) badge.textContent = `${rows.length} manager${rows.length === 1 ? '' : 's'} · ${fmtNum(total)}`;
  }

  function render(model) {
    if (!window.Chart) return;
    const c = APEX.theme.colors();
    renderTrend(model.trend, c);
    renderManager('pendsMgrChart', model.pendsByManager, c, 'Pends', false);
    renderManager('ageingMgrChart', model.ageingByManager, c, 'Ageing', true);
  }

  const onManagerClick = fn => { onBarClick = fn; };
  const resize = () => Object.values(charts).forEach(ch => ch.resize());

  return { render, onManagerClick, resize, helpers: { tooltip, gradient, rgba } };
})();
