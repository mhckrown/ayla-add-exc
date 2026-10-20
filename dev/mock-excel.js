/* Excel simulado (subconjunto de Office.js) para desarrollar y probar el panel sin Excel.
 * No es completo: cubre lo que usa src/excel-bridge.js. Las fórmulas soportan
 * referencias, + - * / ^ &, comparaciones y SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/IF/IFERROR/ROUND/ABS/AND/OR. */
(function () {
  const MAX_ROW = 1048576, MAX_COL = 16384;
  const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];
  const listeners = [];
  let changed = new Set();

  const wb = { name: 'Ventas2026.xlsx', sheets: [], active: 0, sel: { sheet: 0, r1: 1, c1: 1, r2: 1, c2: 1 }, tables: [], names: [], seq: 1 };
  const key = (r, c) => `${r},${c}`;
  const colToNum = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const numToCol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const quote = (n) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  const err = (code, msg) => Object.assign(new Error(msg), { code });

  function newSheet(name) { return { name, cells: new Map(), fmt: new Map(), charts: [], visibility: 'Visible' }; }
  function findSheet(name) { return wb.sheets.find((s) => s.name.toLowerCase() === String(name).toLowerCase()); }

  function parseRef(ref) {
    const s = ref.replace(/\$/g, '').toUpperCase();
    let m;
    if ((m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(s))) {
      const r1 = +m[2], c1 = colToNum(m[1]), r2 = m[4] ? +m[4] : r1, c2 = m[3] ? colToNum(m[3]) : c1;
      return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
    }
    if ((m = /^([A-Z]+):([A-Z]+)$/.exec(s))) return { r1: 1, r2: MAX_ROW, c1: colToNum(m[1]), c2: colToNum(m[2]) };
    if ((m = /^(\d+):(\d+)$/.exec(s))) return { r1: +m[1], r2: +m[2], c1: 1, c2: MAX_COL };
    throw err('InvalidArgument', `Referencia no válida: ${ref}`);
  }
  function a1(r1, c1, r2, c2) {
    if (r1 === 1 && r2 === MAX_ROW) return `${numToCol(c1)}:${numToCol(c2)}`;
    if (c1 === 1 && c2 === MAX_COL) return `${r1}:${r2}`;
    const tl = `${numToCol(c1)}${r1}`;
    return r1 === r2 && c1 === c2 ? tl : `${tl}:${numToCol(c2)}${r2}`;
  }
  function splitSheet(text, defSheet) {
    const i = text.lastIndexOf('!');
    if (i < 0) return { sheet: defSheet, ref: text };
    let n = text.slice(0, i);
    if (n.startsWith("'")) n = n.slice(1, -1).replace(/''/g, "'");
    return { sheet: findSheet(n), ref: text.slice(i + 1) };
  }

  /* ---------- Evaluador de fórmulas ---------- */
  function tokenize(src) {
    const re = /\s*(?:("(?:[^"]|"")*")|(#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A))|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?(?![\w(]))|(\d+(?:\.\d+)?(?:E[+-]?\d+)?%?)|([A-Za-z_][\w.]*)(?=\()|(TRUE|FALSE)\b|(<>|<=|>=|[-+*/^&=<>(),:;]))/giy;
    const out = []; let m; re.lastIndex = 0;
    while (re.lastIndex < src.length && (m = re.exec(src))) {
      if (m[1]) out.push({ t: 'str', v: m[1].slice(1, -1).replace(/""/g, '"') });
      else if (m[2]) out.push({ t: 'err', v: m[2].toUpperCase() });
      else if (m[3]) out.push({ t: 'ref', v: m[3] });
      else if (m[4]) out.push({ t: 'num', v: m[4].endsWith('%') ? parseFloat(m[4]) / 100 : parseFloat(m[4]) });
      else if (m[5]) out.push({ t: 'fn', v: m[5].toUpperCase() });
      else if (m[6]) out.push({ t: 'bool', v: m[6].toUpperCase() === 'TRUE' });
      else out.push({ t: 'op', v: m[7] });
      if (!m[0].length) break;
    }
    if (re.lastIndex < src.length && src.slice(re.lastIndex).trim()) throw new Error('#NAME?');
    return out;
  }

  const isErr = (v) => typeof v === 'string' && ERRORS.includes(v);
  const num = (v) => { if (isErr(v)) throw v; if (v === '' || v === null) return 0; if (typeof v === 'boolean') return v ? 1 : 0; const n = Number(v); if (isNaN(n)) throw '#VALUE!'; return n; };

  function evaluate(formula, sheet, stack) {
    let toks; try { toks = tokenize(formula.slice(1)); } catch { return '#NAME?'; }
    let i = 0;
    const peek = () => toks[i], next = () => toks[i++];
    const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;
    const scalar = (x) => (x && x.range ? (x.cells.length ? x.cells[0] : '') : x);
    function cmp() { let a = concat(); while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = next().v; const b = concat(); const x = scalar(a), y = scalar(b); if (isErr(x)) throw x; if (isErr(y)) throw y; a = { '=': x == y, '<>': x != y, '<': x < y, '>': x > y, '<=': x <= y, '>=': x >= y }[op]; } return a; }
    function concat() { let a = add(); while (isOp('&')) { next(); const b = add(); const x = scalar(a), y = scalar(b); if (isErr(x)) throw x; if (isErr(y)) throw y; a = `${x}${y}`; } return a; }
    function add() { let a = mul(); while (isOp('+') || isOp('-')) { const op = next().v; const b = mul(); a = op === '+' ? num(scalar(a)) + num(scalar(b)) : num(scalar(a)) - num(scalar(b)); } return a; }
    function mul() { let a = pow(); while (isOp('*') || isOp('/')) { const op = next().v; const b = pow(); if (op === '*') a = num(scalar(a)) * num(scalar(b)); else { const d = num(scalar(b)); if (d === 0) throw '#DIV/0!'; a = num(scalar(a)) / d; } } return a; }
    function pow() { let a = unary(); while (isOp('^')) { next(); a = Math.pow(num(scalar(a)), num(scalar(unary()))); } return a; }
    function unary() { if (isOp('-')) { next(); return -num(scalar(unary())); } if (isOp('+')) { next(); return unary(); } return primary(); }
    function primary() {
      const t = next();
      if (!t) throw '#VALUE!';
      if (t.t === 'num' || t.t === 'str' || t.t === 'bool') return t.v;
      if (t.t === 'err') throw t.v;
      if (t.t === 'ref') {
        const { sheet: sh, ref } = splitSheet(t.v, sheet);
        if (!sh) throw '#REF!';
        const b = parseRef(ref);
        const cells = [];
        for (let r = b.r1; r <= Math.min(b.r2, 5000); r++) for (let c = b.c1; c <= Math.min(b.c2, 200); c++) cells.push(cellValue(sh, r, c, stack));
        return { range: true, cells };
      }
      if (t.t === 'op' && t.v === '(') { const v = cmp(); next(); return v; }
      if (t.t === 'fn') {
        next(); // '('
        const thunks = [];
        if (!isOp(')')) {
          while (true) {
            const start = i; let depth = 0;
            while (i < toks.length && !(depth === 0 && (isOp(',') || isOp(';') || isOp(')')))) { if (isOp('(')) depth++; if (isOp(')')) depth--; i++; }
            thunks.push(() => { const save = i; i = start; const v = cmp(); i = save; return v; });
            if (isOp(',') || isOp(';')) { next(); continue; }
            break;
          }
        }
        next(); // ')'
        return callFn(t.v, thunks);
      }
      throw '#VALUE!';
    }
    function flat(thunks) { const out = []; for (const th of thunks) { const v = th(); if (v && v.range) out.push(...v.cells); else out.push(v); } return out; }
    function callFn(name, a) {
      const nums = () => flat(a).filter((v) => { if (isErr(v)) throw v; return typeof v === 'number'; });
      switch (name) {
        case 'SUM': return nums().reduce((s, v) => s + v, 0);
        case 'AVERAGE': { const n = nums(); if (!n.length) throw '#DIV/0!'; return n.reduce((s, v) => s + v, 0) / n.length; }
        case 'MIN': { const n = nums(); return n.length ? Math.min(...n) : 0; }
        case 'MAX': { const n = nums(); return n.length ? Math.max(...n) : 0; }
        case 'COUNT': return nums().length;
        case 'COUNTA': return flat(a).filter((v) => v !== '' && v !== null).length;
        case 'IF': return scalar(a[0]()) ? (a[1] ? scalar(a[1]()) : true) : (a[2] ? scalar(a[2]()) : false);
        case 'IFERROR': try { const v = scalar(a[0]()); if (isErr(v)) throw v; return v; } catch (e) { if (isErr(e)) return scalar(a[1]()); throw e; }
        case 'ROUND': { const f = Math.pow(10, num(scalar(a[1] ? a[1]() : 0))); return Math.round(num(scalar(a[0]())) * f) / f; }
        case 'ABS': return Math.abs(num(scalar(a[0]())));
        case 'AND': return flat(a).every(Boolean);
        case 'OR': return flat(a).some(Boolean);
        default: throw '#NAME?';
      }
    }
    try { const v = scalar(cmp()); return typeof v === 'number' && !isFinite(v) ? '#NUM!' : v; } catch (e) { return isErr(e) ? e : '#VALUE!'; }
  }

  function cellRaw(sheet, r, c) { const v = sheet.cells.get(key(r, c)); return v === undefined ? '' : v; }
  function cellValue(sheet, r, c, stack = new Set()) {
    const raw = cellRaw(sheet, r, c);
    if (typeof raw !== 'string' || !raw.startsWith('=')) return raw;
    const k = `${sheet.name}!${key(r, c)}`;
    if (stack.has(k)) return 0;
    stack.add(k);
    const v = evaluate(raw, sheet, stack);
    stack.delete(k);
    return v;
  }
  function setRaw(sheet, r, c, v) {
    const k = key(r, c);
    if (v === '' || v === null || v === undefined) sheet.cells.delete(k);
    else if (typeof v === 'string' && v !== '' && !v.startsWith('=') && !isNaN(Number(v)) && v.trim() !== '') sheet.cells.set(k, Number(v));
    else sheet.cells.set(k, v);
    changed.add(`${sheet.name}!${k}`);
  }

  function shiftFormula(f, dr, dc) {
    return f.replace(/("(?:[^"]|"")*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]+)(\$?)(\d+)(?![\w(])/g, (m, str, sh, ca, col, ra, row) => {
      if (str) return str;
      const nc = ca ? col : numToCol(Math.max(1, colToNum(col) + dc));
      const nr = ra ? row : Math.max(1, +row + dr);
      return `${sh || ''}${ca}${nc}${ra}${nr}`;
    });
  }

  function usedBounds(sheet) {
    let b = null;
    for (const k of sheet.cells.keys()) {
      const [r, c] = k.split(',').map(Number);
      b = b ? { r1: Math.min(b.r1, r), c1: Math.min(b.c1, c), r2: Math.max(b.r2, r), c2: Math.max(b.c2, c) } : { r1: r, c1: c, r2: r, c2: c };
    }
    return b;
  }

  function refsOf(formula, sheet) {
    const out = [];
    const re = /("(?:[^"]|"")*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?(?![\w(])/g;
    let m;
    while ((m = re.exec(formula))) {
      if (m[1]) continue;
      const { sheet: sh, ref } = splitSheet(m[0], sheet);
      if (!sh) continue;
      const b = parseRef(ref);
      out.push({ sheet: sh, ...b, text: `${quote(sh.name)}!${a1(b.r1, b.c1, b.r2, b.c2)}` });
    }
    return out;
  }

  /* ---------- Objetos de la API ---------- */
  function grid(rows, cols, fn) { return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => fn(r, c))); }

  function makeRange(sheet, b) {
    const { r1, c1, r2, c2 } = b;
    const rows = r2 - r1 + 1, cols = c2 - c1 + 1;
    const each = (fn) => { for (let r = r1; r <= Math.min(r2, r1 + 2000); r++) for (let c = c1; c <= Math.min(c2, c1 + 200); c++) fn(r, c); };
    const write = (v) => {
      if (!Array.isArray(v)) { each((r, c) => setRaw(sheet, r, c, v)); return; }
      if (v.length !== rows || v.some((row) => row.length !== cols)) throw err('InvalidArgument', `Las dimensiones no coinciden con el rango (${rows}x${cols}).`);
      v.forEach((row, i) => row.forEach((x, j) => setRaw(sheet, r1 + i, c1 + j, x)));
    };
    const setFmt = (prop, val) => each((r, c) => { const k = key(r, c); sheet.fmt.set(k, { ...(sheet.fmt.get(k) || {}), [prop]: val }); changed.add(`${sheet.name}!${k}`); });
    const range = {
      isNullObject: false,
      get address() { return `${quote(sheet.name)}!${a1(r1, c1, r2, c2)}`; },
      get rowCount() { return rows; },
      get columnCount() { return cols; },
      get values() { return grid(rows, cols, (r, c) => cellValue(sheet, r1 + r, c1 + c)); },
      set values(v) { write(v); },
      get formulas() { return grid(rows, cols, (r, c) => cellRaw(sheet, r1 + r, c1 + c)); },
      set formulas(v) { write(v); },
      get numberFormat() { return grid(rows, cols, (r, c) => (sheet.fmt.get(key(r1 + r, c1 + c)) || {}).numFmt || 'General'); },
      set numberFormat(v) { v.forEach((row, i) => row.forEach((x, j) => { const k = key(r1 + i, c1 + j); sheet.fmt.set(k, { ...(sheet.fmt.get(k) || {}), numFmt: x }); })); },
      get worksheet() { return wsApi(sheet); },
      load() { return range; },
      getCell(r, c) { return makeRange(sheet, { r1: r1 + r, c1: c1 + c, r2: r1 + r, c2: c1 + c }); },
      getResizedRange(dr, dc) { return makeRange(sheet, { r1, c1, r2: r2 + dr, c2: c2 + dc }); },
      select() { wb.active = wb.sheets.indexOf(sheet); wb.sel = { sheet: wb.active, r1, c1, r2: Math.min(r2, 200), c2: Math.min(c2, 50) }; },
      clear(applyTo = 'All') { each((r, c) => { if (applyTo !== 'Formats') setRaw(sheet, r, c, ''); if (applyTo !== 'Contents') sheet.fmt.delete(key(r, c)); }); },
      merge() {}, unmerge() {},
      insert(shift) { shiftCells(sheet, b, shift === 'Right' ? 'cols' : 'rows', +1); },
      delete(shift) { shiftCells(sheet, b, shift === 'Left' ? 'cols' : 'rows', -1); },
      autoFill(dest) {
        const d = typeof dest === 'string' ? parseRef(splitSheet(dest).ref) : dest._b;
        for (let r = d.r1; r <= d.r2; r++) for (let c = d.c1; c <= d.c2; c++) {
          if (r >= r1 && r <= r2 && c >= c1 && c <= c2) continue;
          const sr = r1 + ((r - r1) % rows + rows) % rows, sc = c1 + ((c - c1) % cols + cols) % cols;
          const raw = cellRaw(sheet, sr, sc);
          setRaw(sheet, r, c, typeof raw === 'string' && raw.startsWith('=') ? shiftFormula(raw, r - sr, c - sc) : raw);
        }
      },
      format: {
        font: { set bold(v) { setFmt('bold', v); }, set italic(v) { setFmt('italic', v); }, set color(v) { setFmt('color', v); }, set size(v) { setFmt('size', v); } },
        fill: { set color(v) { setFmt('fill', v); }, clear() { setFmt('fill', null); } },
        set horizontalAlignment(v) { setFmt('align', v); }, set wrapText(v) {}, set columnWidth(v) {}, set rowHeight(v) {},
        autofitColumns() {}, autofitRows() {},
        borders: { getItem: () => ({ style: '', color: '' }) },
      },
      conditionalFormats: { add: (type) => ({ type, colorScale: {}, dataBar: { positiveFormat: {} }, cellValue: { format: { fill: {}, font: {} } } }) },
      sort: {
        apply(fields, matchCase, hasHeaders) {
          const start = hasHeaders ? r1 + 1 : r1;
          const rowsData = [];
          for (let r = start; r <= r2; r++) rowsData.push({ raw: grid(1, cols, (_, c) => cellRaw(sheet, r, c1 + c))[0], val: grid(1, cols, (_, c) => cellValue(sheet, r, c1 + c))[0] });
          const f = fields[0];
          rowsData.sort((a, b2) => { const x = a.val[f.key], y = b2.val[f.key]; const res = x < y ? -1 : x > y ? 1 : 0; return f.ascending === false ? -res : res; });
          rowsData.forEach((row, i) => row.raw.forEach((v, j) => setRaw(sheet, start + i, c1 + j, v)));
        },
      },
      getDirectPrecedents() {
        const raw = cellRaw(sheet, r1, c1);
        if (typeof raw !== 'string' || !raw.startsWith('=')) throw err('ItemNotFound', 'Sin precedentes');
        return { addresses: refsOf(raw, sheet).map((x) => x.text), load() { return this; } };
      },
      getDirectDependents() {
        const out = [];
        for (const sh of wb.sheets) for (const [k, raw] of sh.cells) {
          if (typeof raw !== 'string' || !raw.startsWith('=')) continue;
          if (refsOf(raw, sh).some((x) => x.sheet === sheet && r1 >= x.r1 && r1 <= x.r2 && c1 >= x.c1 && c1 <= x.c2)) {
            const [r, c] = k.split(',').map(Number);
            out.push(`${quote(sh.name)}!${a1(r, c, r, c)}`);
          }
        }
        if (!out.length) throw err('ItemNotFound', 'Sin dependientes');
        return { addresses: out, load() { return this; } };
      },
      _b: b,
      _sheet: sheet,
    };
    return range;
  }

  function shiftCells(sheet, b, axis, dir) {
    const count = axis === 'rows' ? b.r2 - b.r1 + 1 : b.c2 - b.c1 + 1;
    const startAt = axis === 'rows' ? b.r1 : b.c1;
    const move = (map) => {
      const entries = [...map.entries()];
      map.clear();
      for (const [k, v] of entries) {
        let [r, c] = k.split(',').map(Number);
        const p = axis === 'rows' ? r : c;
        if (dir < 0 && p >= startAt && p < startAt + count) continue;
        if (p >= startAt) { if (axis === 'rows') r += dir * count; else c += dir * count; }
        map.set(key(r, c), v);
      }
    };
    move(sheet.cells); move(sheet.fmt);
    changed.add('*');
  }

  function wsApi(sheet) {
    const api = {
      isNullObject: false,
      get name() { return sheet.name; },
      set name(v) { if (findSheet(v) && findSheet(v) !== sheet) throw err('InvalidArgument', 'Ya existe una hoja con ese nombre'); sheet.name = v; changed.add('*'); },
      get position() { return wb.sheets.indexOf(sheet); },
      get visibility() { return sheet.visibility; },
      load() { return api; },
      getRange(ref) { return makeRange(sheet, parseRef(ref || 'A1')); },
      getUsedRange() { const b = usedBounds(sheet); return makeRange(sheet, b || { r1: 1, c1: 1, r2: 1, c2: 1 }); },
      getUsedRangeOrNullObject() { const b = usedBounds(sheet); return b ? makeRange(sheet, b) : { isNullObject: true, load() { return this; } }; },
      activate() { wb.active = wb.sheets.indexOf(sheet); wb.sel = { sheet: wb.active, r1: 1, c1: 1, r2: 1, c2: 1 }; changed.add('*'); },
      delete() { if (wb.sheets.length === 1) throw err('InvalidOperation', 'No se puede eliminar la única hoja'); wb.sheets.splice(wb.sheets.indexOf(sheet), 1); wb.active = 0; wb.sel = { sheet: 0, r1: 1, c1: 1, r2: 1, c2: 1 }; changed.add('*'); },
      protection: { protect() { sheet.protected = true; } },
      autoFilter: { apply() {} },
      charts: {
        get items() { return sheet.charts; },
        load() { return this; },
        add(type, range) {
          const chart = { name: `Gráfico ${wb.seq++}`, type, source: range.address, title: { text: '' }, setPosition(cell) { this.position = String(cell); }, load() { return this; }, delete() { sheet.charts.splice(sheet.charts.indexOf(chart), 1); changed.add('*'); } };
          sheet.charts.push(chart); changed.add('*');
          return chart;
        },
        getItem(name) { const c = sheet.charts.find((x) => x.name === name); if (!c) throw err('ItemNotFound', `No existe el gráfico ${name}`); return c; },
      },
    };
    return api;
  }

  const worksheets = {
    get items() { return wb.sheets.map(wsApi); },
    load() { return worksheets; },
    getActiveWorksheet() { return wsApi(wb.sheets[wb.active]); },
    getItem(name) { const s = findSheet(name); if (!s) throw err('ItemNotFound', `No existe la hoja "${name}"`); return wsApi(s); },
    getItemOrNullObject(name) { const s = findSheet(name); return s ? wsApi(s) : { isNullObject: true, load() { return this; } }; },
    add(name) { let n = name; if (!n) { let i = wb.sheets.length + 1; while (findSheet(`Hoja${i}`)) i++; n = `Hoja${i}`; } if (findSheet(n)) throw err('InvalidArgument', 'Ya existe una hoja con ese nombre'); const s = newSheet(n); wb.sheets.push(s); changed.add('*'); return wsApi(s); },
  };

  function resolve(rangeOrAddress) {
    if (typeof rangeOrAddress !== 'string') return rangeOrAddress;
    const { sheet, ref } = splitSheet(rangeOrAddress, wb.sheets[wb.active]);
    return makeRange(sheet, parseRef(ref));
  }

  const tables = {
    get items() { return wb.tables; },
    load() { return tables; },
    add(rangeOrAddress, hasHeaders) {
      const range = resolve(rangeOrAddress);
      const t = {
        name: `Tabla${wb.seq++}`, style: 'TableStyleMedium2', showHeaders: hasHeaders, _sheet: range._sheet, _b: range._b,
        getRange: () => makeRange(t._sheet, t._b), load() { return t; },
        convertToRange() { wb.tables.splice(wb.tables.indexOf(t), 1); changed.add('*'); return t.getRange(); },
        delete() { t.convertToRange(); },
      };
      wb.tables.push(t); changed.add('*');
      return t;
    },
    getItem(name) { const t = wb.tables.find((x) => x.name === name); if (!t) throw err('ItemNotFound', `No existe la tabla ${name}`); return t; },
  };

  const names = {
    get items() { return wb.names; },
    load() { return names; },
    add(name, range) { const n = { name, formula: `=${range.address}`, delete() { wb.names.splice(wb.names.indexOf(n), 1); } }; wb.names.push(n); return n; },
    getItem(name) { const n = wb.names.find((x) => x.name === name); if (!n) throw err('ItemNotFound', `No existe el nombre ${name}`); return n; },
  };

  const workbook = {
    get name() { return wb.name; },
    load() { return workbook; },
    worksheets, tables, names,
    getSelectedRange() { const s = wb.sel; return makeRange(wb.sheets[s.sheet] || wb.sheets[0], s); },
  };

  function notify() {
    if (!changed.size) return;
    const c = changed; changed = new Set();
    listeners.forEach((fn) => fn(c));
  }

  const Excel = {
    async run(cb) {
      const ctx = { workbook, async sync() { await new Promise((r) => setTimeout(r, 15)); notify(); } };
      try { return await cb(ctx); } finally { notify(); }
    },
  };
  const Office = {
    HostType: { Excel: 'Excel' },
    context: { requirements: { isSetSupported: () => true }, roamingSettings: { get: () => null, set() {}, saveAsync() {} } },
    onReady(cb) { const info = { host: 'Excel', platform: 'PC' }; setTimeout(() => cb && cb(info), 0); return Promise.resolve(info); },
  };

  /* ---------- Datos de ejemplo ---------- */
  const ventas = newSheet('Ventas'); const sup = newSheet('Supuestos'); const res = newSheet('Resumen');
  wb.sheets.push(ventas, sup, res);
  const put = (sh, rows, r0 = 1) => rows.forEach((row, i) => row.forEach((v, j) => { if (v !== '' && v !== null) sh.cells.set(key(r0 + i, j + 1), v); }));
  const meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const regiones = ['Norte', 'Sur', 'Centro'];
  put(ventas, [['Mes', 'Región', 'Unidades', 'Precio', 'Total']]);
  meses.forEach((m, i) => put(ventas, [[m, regiones[i % 3], 120 + ((i * 37) % 90), 25 + (i % 4) * 5, `=C${i + 2}*D${i + 2}`]], i + 2));
  put(ventas, [['Total', '', '=SUM(C2:C13)', '', '=SUM(E2:E13)']], 14);
  put(sup, [['Supuesto', 'Valor'], ['Tasa IVA', 0.16], ['Meta anual', 50000], ['Crecimiento', 0.08]]);
  put(res, [['Indicador', 'Valor'], ['Ventas totales', '=Ventas!E14'], ['IVA', '=B2*Supuestos!B2'], ['Avance vs meta', '=B2/Supuestos!B3'], ['Proyección próximo año', '=B2*(1+Supuestos!B4)'], ['Ticket promedio', '=B2/#REF!'], ['Precio medio', '=AVERAGE(Ventas!D2:D13)/Supuestos!B9']]);
  wb.sel = { sheet: 0, r1: 1, c1: 1, r2: 14, c2: 5 };

  window.MockExcel = {
    Excel, Office, wb, cellValue, cellRaw, setRaw, numToCol, key, notify,
    onChange(fn) { listeners.push(fn); },
    select(sheetIdx, r1, c1, r2 = r1, c2 = c1) { wb.active = sheetIdx; wb.sel = { sheet: sheetIdx, r1, c1, r2, c2 }; },
  };
})();
