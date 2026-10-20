/* Cuadrícula visual del Excel simulado. */
(function () {
  const M = window.MockExcel;
  const gridEl = document.getElementById('grid');
  const tabsEl = document.getElementById('tabs');
  const addrEl = document.getElementById('addr');
  const fxEl = document.getElementById('fx');
  const logEl = document.getElementById('log');
  document.getElementById('wbname').textContent = ' — ' + M.wb.name;

  function fmtValue(v, fmt) {
    if (typeof v === 'number') {
      if (fmt && fmt.includes('%')) return (v * 100).toFixed(fmt.includes('.0') ? 1 : 0) + '%';
      if (fmt && /[$€]/.test(fmt)) return (fmt.includes('€') ? '€' : '$') + v.toLocaleString('es', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      return Number.isInteger(v) ? String(v) : v.toLocaleString('es', { maximumFractionDigits: 4 });
    }
    if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
    return String(v);
  }

  function render(changed = new Set()) {
    const sheet = M.wb.sheets[M.wb.active];
    let maxR = 30, maxC = 10;
    for (const k of sheet.cells.keys()) { const [r, c] = k.split(',').map(Number); maxR = Math.max(maxR, r + 5); maxC = Math.max(maxC, c + 2); }
    const s = M.wb.sel.sheet === M.wb.active ? M.wb.sel : null;
    const tableCells = new Set();
    M.wb.tables.filter((t) => t._sheet === sheet).forEach((t) => { for (let r = t._b.r1; r <= t._b.r2; r++) for (let c = t._b.c1; c <= t._b.c2; c++) tableCells.add(M.key(r, c) + (r === t._b.r1 ? 'h' : '')); });
    let html = '<tr><th></th>' + Array.from({ length: maxC }, (_, i) => `<th>${M.numToCol(i + 1)}</th>`).join('') + '</tr>';
    for (let r = 1; r <= maxR; r++) {
      html += `<tr><td class="rh">${r}</td>`;
      for (let c = 1; c <= maxC; c++) {
        const k = M.key(r, c);
        const v = M.cellValue(sheet, r, c);
        const f = sheet.fmt.get(k) || {};
        const cls = [];
        if (typeof v === 'number') cls.push('num');
        if (typeof v === 'string' && v.startsWith('#')) cls.push('err');
        if (s && r >= s.r1 && r <= s.r2 && c >= s.c1 && c <= s.c2) cls.push('sel');
        if (changed.has(`${sheet.name}!${k}`)) cls.push('flash');
        let style = '';
        if (f.bold) style += 'font-weight:700;';
        if (f.italic) style += 'font-style:italic;';
        if (f.color) style += `color:${f.color};`;
        if (f.fill) style += `background:${f.fill};`;
        else if (tableCells.has(k + 'h')) style += 'background:#4472c4;color:#fff;font-weight:600;';
        else if (tableCells.has(k)) style += r % 2 ? 'background:#d9e1f2;' : 'background:#eef2fa;';
        if (f.align) style += `text-align:${f.align.toLowerCase()};`;
        html += `<td data-r="${r}" data-c="${c}" class="${cls.join(' ')}" style="${style}">${fmtValue(v, f.numFmt).replace(/</g, '&lt;')}</td>`;
      }
      html += '</tr>';
    }
    gridEl.innerHTML = html;
    tabsEl.innerHTML = M.wb.sheets.map((sh, i) => `<button data-i="${i}" class="${i === M.wb.active ? 'active' : ''}">${sh.name}</button>`).join('');
    const charts = sheet.charts.map((c) => `${c.name} (${c.type}, ${c.source}${c.title.text ? ', "' + c.title.text + '"' : ''})`);
    if (charts.length) tabsEl.innerHTML += `<span style="margin-left:auto;padding:5px;color:#217346">📊 ${charts.join(' · ')}</span>`;
    updateFormulaBar();
  }

  function updateFormulaBar() {
    const s = M.wb.sel;
    const sheet = M.wb.sheets[s.sheet];
    if (!sheet) return;
    addrEl.textContent = `${M.numToCol(s.c1)}${s.r1}` + (s.r2 !== s.r1 || s.c2 !== s.c1 ? `:${M.numToCol(s.c2)}${s.r2}` : '');
    if (document.activeElement !== fxEl) fxEl.value = M.cellRaw(sheet, s.r1, s.c1);
  }

  let dragStart = null;
  gridEl.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td[data-r]');
    if (!td) return;
    dragStart = { r: +td.dataset.r, c: +td.dataset.c };
    M.select(M.wb.active, dragStart.r, dragStart.c);
    render();
  });
  gridEl.addEventListener('mouseover', (e) => {
    const td = e.target.closest('td[data-r]');
    if (!dragStart || !td || e.buttons !== 1) return;
    const r = +td.dataset.r, c = +td.dataset.c;
    M.select(M.wb.active, Math.min(r, dragStart.r), Math.min(c, dragStart.c), Math.max(r, dragStart.r), Math.max(c, dragStart.c));
    render();
  });
  document.addEventListener('mouseup', () => { dragStart = null; });
  tabsEl.addEventListener('click', (e) => {
    const b = e.target.closest('button[data-i]');
    if (!b) return;
    M.select(+b.dataset.i, 1, 1);
    render();
  });
  fxEl.addEventListener('keydown', (e) => {
    if (e.key !== 'Enter') return;
    const s = M.wb.sel;
    M.setRaw(M.wb.sheets[s.sheet], s.r1, s.c1, fxEl.value);
    M.notify();
    fxEl.blur();
  });

  M.onChange((changed) => {
    render(changed);
    const n = [...changed].filter((k) => k !== '*').length;
    const line = document.createElement('div');
    line.textContent = `${new Date().toLocaleTimeString()} — ${n ? n + ' celda(s) modificadas' : 'cambio estructural'}`;
    logEl.prepend(line);
  });

  render();
})();
