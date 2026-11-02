/* APEX Dashboard - KPI cards (Containers #1 and #3) */
window.APEX = window.APEX || {};

APEX.kpi = (() => {
  const { $, animateCount, fmtNum } = APEX.utils;

  function renderCard(cardId, s) {
    const card = $(`#${cardId}`);
    if (!card) return;
    animateCount($('[data-kpi="total"]', card), s.total, 1400);
    ['completed', 'pending', 'blank'].forEach(k => {
      animateCount($(`[data-kpi="${k}"]`, card), s[k]);
      const bar = $(`[data-bar="${k}"]`, card);
      if (bar) bar.style.width = s.total ? `${(s[k] / s.total) * 100}%` : '0%';
    });
    const note = $('[data-kpi-note]', card);
    if (note) {
      const parts = [];
      if (!s.total) parts.push('No records found for this category.');
      else parts.push(`Pending rate ${((s.pending / s.total) * 100).toFixed(1)}% · Completion rate ${((s.completed / s.total) * 100).toFixed(1)}%`);
      if (s.other) parts.push(`${fmtNum(s.other)} with other status values`);
      note.textContent = parts.join(' · ');
    }
  }

  function render(model) {
    renderCard('kpi-pends', model.pends);
    renderCard('kpi-ageing', model.ageing);
  }

  return { render };
})();
