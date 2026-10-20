/* === Ayla Add-Exc — Excel Bridge === */
/* Herramientas que el modelo de IA ejecuta en Excel vía Office.js */

const MAX_READ_CELLS = 5000;
const MAX_SCAN_CELLS = 200000;
const MAX_UNDO_CELLS = 50000;
const EXCEL_ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!'];

/* --- Direcciones --- */

function splitAddress(address) {
  const s = String(address || '').trim();
  const i = s.lastIndexOf('!');
  if (i === -1) return { sheet: null, ref: s };
  let sheet = s.slice(0, i);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, ref: s.slice(i + 1) };
}

function quoteSheet(name) {
  return /^[A-Za-z_][\w.]*$/.test(name) ? name : `'${String(name).replace(/'/g, "''")}'`;
}

function colToNum(col) {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Celda superior izquierda de una referencia ("B3:D9" → {row:3, col:2}) */
function topLeft(ref) {
  const m = /\$?([A-Z]{1,3})?\$?(\d+)?/i.exec(splitAddress(ref).ref.split(':')[0]) || [];
  return { col: m[1] ? colToNum(m[1]) : 1, row: m[2] ? parseInt(m[2], 10) : 1 };
}

function getSheet(ctx, name) {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

/** Acepta "A1:B2", "Hoja1!A1:B2" o "'Mi hoja'!A1" (+ sheetName opcional) */
function getRange(ctx, address, sheetName) {
  const { sheet, ref } = splitAddress(address);
  if (!ref) throw new Error('Falta la dirección del rango.');
  return getSheet(ctx, sheet || sheetName).getRange(ref);
}

function isSetSupported(version) {
  try {
    return !!(Office && Office.context && Office.context.requirements &&
      Office.context.requirements.isSetSupported('ExcelApi', version));
  } catch {
    return false;
  }
}

/* --- Registro de cambios (para Deshacer) --- */

const ChangeTracker = {
  current: null,
  begin() { this.current = []; },
  end() { const c = this.current || []; this.current = null; return c; },
  record(change) {
    if (!this.current) return;
    // Un cambio estructural desplaza celdas: las instantáneas previas ya no son seguras de restaurar.
    if (change.structural) this.current.forEach((c) => { c.undo = null; });
    this.current.push(change);
  },
};

/** Guarda fórmulas/valores de un rango antes de modificarlo para poder deshacer. */
async function snapshotRange(ctx, range, label) {
  range.load('address, rowCount, columnCount');
  await ctx.sync();
  const address = range.address;
  if (range.rowCount * range.columnCount > MAX_UNDO_CELLS) {
    ChangeTracker.record({ label, address, undo: null });
    return address;
  }
  range.load('formulas');
  await ctx.sync();
  const saved = range.formulas;
  ChangeTracker.record({
    label,
    address,
    undo: () => Excel.run(async (c) => {
      getRange(c, address).formulas = saved;
      await c.sync();
    }),
  });
  return address;
}

function recordNonUndoable(label, address, structural = false) {
  ChangeTracker.record({ label, address, undo: null, structural });
}

/* --- Utilidades de lectura --- */

/** Fórmulas como mapa disperso { "B5": "=SUM(B1:B4)" } */
function sparseFormulas(formulas, startRow, startCol) {
  const out = {};
  let count = 0;
  (formulas || []).forEach((row, r) => row.forEach((f, c) => {
    if (typeof f === 'string' && f.startsWith('=') && count < 2000) {
      out[`${numToCol(startCol + c)}${startRow + r}`] = f;
      count++;
    }
  }));
  return out;
}

async function readRangeData(ctx, range, { includeFormulas = true, maxCells = MAX_READ_CELLS } = {}) {
  range.load('address, rowCount, columnCount');
  await ctx.sync();
  let target = range;
  const total = range.rowCount * range.columnCount;
  let truncated = false;
  if (total > maxCells) {
    const cols = Math.min(range.columnCount, maxCells);
    const rows = Math.max(1, Math.min(range.rowCount, Math.floor(maxCells / cols)));
    target = range.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    target.load('address');
    truncated = true;
  }
  target.load(includeFormulas ? 'values, formulas' : 'values');
  await ctx.sync();
  const start = topLeft(range.address);
  const result = {
    address: range.address,
    startCell: `${numToCol(start.col)}${start.row}`,
    rows: range.rowCount,
    cols: range.columnCount,
    values: target.values,
  };
  if (includeFormulas) result.formulas = sparseFormulas(target.formulas, start.row, start.col);
  if (truncated) {
    result.truncated = true;
    result.returnedAddress = target.address;
    result.note = `Rango grande: solo se devolvieron ${target.values.length} filas. Lee el resto por partes con read_range.`;
  }
  return result;
}

async function usedRangeOf(ctx, sheet) {
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load('address, rowCount, columnCount');
  await ctx.sync();
  return used.isNullObject ? null : used;
}

async function sheetsToScan(ctx, sheetName) {
  if (sheetName) {
    const s = ctx.workbook.worksheets.getItem(sheetName);
    s.load('name');
    await ctx.sync();
    return [s];
  }
  const sheets = ctx.workbook.worksheets;
  sheets.load('items/name');
  await ctx.sync();
  return sheets.items;
}

/** Extrae referencias de una fórmula (fallback si no hay getDirectPrecedents). */
function referencesInFormula(formula) {
  const refs = new Set();
  const re = /(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g;
  const withoutStrings = String(formula || '').replace(/"[^"]*"/g, '""');
  let m;
  while ((m = re.exec(withoutStrings))) refs.add(m[0].replace(/\$/g, ''));
  return [...refs];
}

/* ===================================================================== */
/* Definición de herramientas                                             */
/* ===================================================================== */

const ADDRESS_DESC = 'Dirección A1. Puede incluir hoja: "B2:D10", "Ventas!A1:C5" o "\'Mi hoja\'!A1".';
const SHEET_PARAM = { type: 'string', description: 'Nombre de la hoja (opcional; por defecto la hoja de la dirección o la activa).' };

const TOOLS = [
  /* ---------- Lectura ---------- */
  {
    name: 'get_workbook_overview',
    description: 'Resumen de todo el libro: hojas con su rango usado, encabezados, tablas, gráficos, nombres definidos, hoja activa y selección. Úsalo para orientarte antes de trabajar.',
    parameters: { type: 'object', properties: {}, required: [] },
    label: () => 'Revisando la estructura del libro',
    async run() {
      return Excel.run(async (ctx) => {
        const wb = ctx.workbook;
        wb.load('name');
        const sheets = wb.worksheets;
        sheets.load('items/name, items/position, items/visibility');
        const active = wb.worksheets.getActiveWorksheet();
        active.load('name');
        const sel = wb.getSelectedRange();
        sel.load('address');
        const tables = wb.tables;
        tables.load('items/name');
        const names = wb.names;
        names.load('items/name, items/formula');
        await ctx.sync();

        const info = sheets.items.map((s) => {
          const used = s.getUsedRangeOrNullObject(true);
          used.load('address, rowCount, columnCount');
          s.charts.load('items/name');
          return { s, used };
        });
        const tableRanges = tables.items.map((t) => {
          const r = t.getRange();
          r.load('address');
          return { t, r };
        });
        await ctx.sync();

        const previews = info.map(({ used }) => {
          if (used.isNullObject) return null;
          const cols = Math.min(used.columnCount, 15);
          const rows = Math.min(used.rowCount, 3);
          const p = used.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
          p.load('values');
          return p;
        });
        await ctx.sync();

        return {
          workbook: wb.name,
          activeSheet: active.name,
          selection: sel.address,
          sheets: info.map(({ s, used }, i) => ({
            name: s.name,
            visible: s.visibility === 'Visible',
            usedRange: used.isNullObject ? null : used.address,
            rows: used.isNullObject ? 0 : used.rowCount,
            cols: used.isNullObject ? 0 : used.columnCount,
            firstRows: previews[i] ? previews[i].values : [],
            charts: s.charts.items.map((c) => c.name),
          })),
          tables: tableRanges.map(({ t, r }) => ({ name: t.name, address: r.address })),
          namedRanges: names.items.map((n) => ({ name: n.name, refersTo: n.formula })),
        };
      });
    },
  },
  {
    name: 'read_range',
    description: 'Lee valores y fórmulas de un rango (de cualquier hoja). Las fórmulas se devuelven como mapa {celda: fórmula}. Rangos grandes se truncan; léelos por partes.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        includeFormulas: { type: 'boolean', description: 'Incluir fórmulas (default: true)' },
      },
      required: ['address'],
    },
    label: (a) => `Leyendo ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => readRangeData(ctx, getRange(ctx, args.address, args.sheetName), {
        includeFormulas: args.includeFormulas !== false,
      }));
    },
  },
  {
    name: 'get_selected_range',
    description: 'Lee el rango seleccionado actualmente por el usuario (valores y fórmulas).',
    parameters: { type: 'object', properties: {}, required: [] },
    label: () => 'Leyendo la selección',
    async run() {
      return Excel.run(async (ctx) => readRangeData(ctx, ctx.workbook.getSelectedRange()));
    },
  },
  {
    name: 'get_worksheet_data',
    description: 'Lee el rango usado completo de una hoja (por defecto la activa).',
    parameters: { type: 'object', properties: { sheetName: SHEET_PARAM }, required: [] },
    label: (a) => `Leyendo la hoja ${a.sheetName || 'activa'}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const sheet = getSheet(ctx, args.sheetName);
        sheet.load('name');
        const used = await usedRangeOf(ctx, sheet);
        if (!used) return { sheet: sheet.name, empty: true, values: [] };
        return { sheet: sheet.name, ...(await readRangeData(ctx, used)) };
      });
    },
  },
  {
    name: 'find_in_workbook',
    description: 'Busca un texto o número en las celdas del libro (o de una hoja). Devuelve las celdas que coinciden.',
    parameters: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Texto a buscar' },
        sheetName: SHEET_PARAM,
        matchCase: { type: 'boolean' },
        exact: { type: 'boolean', description: 'Coincidencia exacta de la celda completa' },
      },
      required: ['query'],
    },
    label: (a) => `Buscando "${a.query}"`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const q = args.matchCase ? String(args.query) : String(args.query).toLowerCase();
        const matches = [];
        let scanned = 0;
        for (const sheet of await sheetsToScan(ctx, args.sheetName)) {
          const used = await usedRangeOf(ctx, sheet);
          if (!used) continue;
          if (scanned + used.rowCount * used.columnCount > MAX_SCAN_CELLS) break;
          scanned += used.rowCount * used.columnCount;
          used.load('values');
          await ctx.sync();
          const start = topLeft(used.address);
          used.values.forEach((row, r) => row.forEach((v, c) => {
            if (matches.length >= 100 || v === '' || v === null) return;
            const text = args.matchCase ? String(v) : String(v).toLowerCase();
            if (args.exact ? text === q : text.includes(q)) {
              matches.push({ cell: `${quoteSheet(sheet.name)}!${numToCol(start.col + c)}${start.row + r}`, value: v });
            }
          }));
        }
        return { query: args.query, count: matches.length, matches, limited: matches.length >= 100 };
      });
    },
  },
  {
    name: 'find_errors',
    description: 'Encuentra celdas con errores (#REF!, #DIV/0!, #N/A, #VALUE!, #NAME?…) y sus fórmulas, en todo el libro o una hoja.',
    parameters: { type: 'object', properties: { sheetName: SHEET_PARAM }, required: [] },
    label: () => 'Buscando errores en fórmulas',
    async run(args) {
      return Excel.run(async (ctx) => {
        const errors = [];
        for (const sheet of await sheetsToScan(ctx, args.sheetName)) {
          const used = await usedRangeOf(ctx, sheet);
          if (!used || used.rowCount * used.columnCount > MAX_SCAN_CELLS) continue;
          used.load('values, formulas');
          await ctx.sync();
          const start = topLeft(used.address);
          used.values.forEach((row, r) => row.forEach((v, c) => {
            if (errors.length < 200 && typeof v === 'string' && EXCEL_ERRORS.includes(v)) {
              errors.push({
                cell: `${quoteSheet(sheet.name)}!${numToCol(start.col + c)}${start.row + r}`,
                error: v,
                formula: used.formulas[r][c],
              });
            }
          }));
        }
        return { count: errors.length, errors };
      });
    },
  },
  {
    name: 'get_cell_dependencies',
    description: 'Explica una celda: su fórmula, valor, las celdas de las que depende (precedentes) y las que dependen de ella (dependientes). Úsalo para explicar o auditar fórmulas.',
    parameters: { type: 'object', properties: { address: { type: 'string', description: 'Celda, ej: "Resumen!C10"' }, sheetName: SHEET_PARAM }, required: ['address'] },
    label: (a) => `Analizando dependencias de ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const cell = getRange(ctx, args.address, args.sheetName).getCell(0, 0);
        cell.load('address, formulas, values, numberFormat');
        await ctx.sync();
        const formula = cell.formulas[0][0];
        const out = { address: cell.address, value: cell.values[0][0], formula, numberFormat: cell.numberFormat[0][0] };
        const hasFormula = typeof formula === 'string' && formula.startsWith('=');
        if (hasFormula) {
          out.precedents = referencesInFormula(formula);
          if (isSetSupported('1.12')) {
            try {
              const p = cell.getDirectPrecedents();
              p.load('addresses');
              await ctx.sync();
              out.precedents = p.addresses;
            } catch { /* sin precedentes o no soportado */ }
          }
        }
        if (isSetSupported('1.15')) {
          try {
            const d = cell.getDirectDependents();
            d.load('addresses');
            await ctx.sync();
            out.dependents = d.addresses;
          } catch {
            out.dependents = [];
          }
        }
        return out;
      });
    },
  },

  /* ---------- Escritura ---------- */
  {
    name: 'set_range_values',
    description: 'Escribe un array 2D de valores y/o fórmulas (cadenas que empiezan con "="). Si la dirección es una sola celda, el rango se expande al tamaño del array. Fórmulas en inglés con coma como separador (SUM, AVERAGE, IF, VLOOKUP…).',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC + ' Celda inicial o rango completo.' },
        sheetName: SHEET_PARAM,
        values: { type: 'array', items: { type: 'array', items: {} }, description: 'Array 2D: [["Mes","Total"],["Ene",100],["Feb","=B2*1.1"]]' },
      },
      required: ['address', 'values'],
    },
    label: (a) => `Escribiendo en ${a.address}`,
    async run(args) {
      const values = Array.isArray(args.values) && Array.isArray(args.values[0]) ? args.values : [[].concat(args.values)];
      const width = Math.max(...values.map((r) => r.length));
      const rows = values.map((r) => [...r, ...Array(width - r.length).fill('')]
        .map((v) => (v === null || v === undefined ? '' : v)));
      return Excel.run(async (ctx) => {
        let range = getRange(ctx, args.address, args.sheetName);
        range.load('rowCount, columnCount');
        await ctx.sync();
        if (range.rowCount === 1 && range.columnCount === 1) {
          range = range.getResizedRange(rows.length - 1, width - 1);
        } else if (range.rowCount !== rows.length || range.columnCount !== width) {
          throw new Error(`El rango tiene ${range.rowCount}x${range.columnCount} celdas pero los datos son ${rows.length}x${width}.`);
        }
        const address = await snapshotRange(ctx, range, 'Valores escritos');
        range.formulas = rows;
        await ctx.sync();
        return { address, rows: rows.length, cols: width };
      });
    },
  },
  {
    name: 'set_cell_value',
    description: 'Escribe un valor o fórmula en una celda (o el mismo valor en todas las celdas de un rango).',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        value: { type: ['string', 'number', 'boolean'], description: 'Valor o fórmula (ej: "=SUM(A1:A10)")' },
      },
      required: ['address', 'value'],
    },
    label: (a) => `Escribiendo en ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        range.load('rowCount, columnCount');
        await ctx.sync();
        const address = await snapshotRange(ctx, range, 'Valor escrito');
        range.formulas = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(args.value ?? ''));
        await ctx.sync();
        return { address, written: true };
      });
    },
  },
  {
    name: 'add_formula',
    description: 'Escribe una fórmula en una celda o rango. Las referencias relativas se ajustan en cada celda del rango (igual que arrastrar). Fórmulas en inglés (SUM, AVERAGE, IF…), separador coma.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        formula: { type: 'string', description: 'Fórmula para la primera celda, ej: "=B2*C2"' },
      },
      required: ['address', 'formula'],
    },
    label: (a) => `Añadiendo fórmula en ${a.address}`,
    async run(args) {
      const formula = String(args.formula).startsWith('=') ? args.formula : `=${args.formula}`;
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        range.load('rowCount, columnCount');
        await ctx.sync();
        const address = await snapshotRange(ctx, range, 'Fórmula añadida');
        const first = range.getCell(0, 0);
        first.formulas = [[formula]];
        if (range.rowCount > 1 || range.columnCount > 1) first.autoFill(range, 'FillDefault');
        await ctx.sync();
        return { address, formula };
      });
    },
  },
  {
    name: 'fill_range',
    description: 'Autorrellena desde un rango origen hasta un rango destino (como arrastrar el controlador de relleno). El destino debe incluir el origen.',
    parameters: {
      type: 'object',
      properties: {
        sourceAddress: { type: 'string', description: 'Rango origen, ej: "D2"' },
        destinationAddress: { type: 'string', description: 'Rango destino que incluye el origen, ej: "D2:D50"' },
        sheetName: SHEET_PARAM,
      },
      required: ['sourceAddress', 'destinationAddress'],
    },
    label: (a) => `Autorrellenando ${a.destinationAddress}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const source = getRange(ctx, args.sourceAddress, args.sheetName);
        const dest = getRange(ctx, args.destinationAddress, args.sheetName);
        const address = await snapshotRange(ctx, dest, 'Autorrelleno');
        source.autoFill(dest, 'FillDefault');
        await ctx.sync();
        return { address, filled: true };
      });
    },
  },
  {
    name: 'clear_range',
    description: 'Borra el contenido, el formato o todo de un rango.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        applyTo: { type: 'string', enum: ['Contents', 'Formats', 'All'], description: 'Qué borrar (default: Contents)' },
      },
      required: ['address'],
    },
    confirm: (a) => (a.applyTo === 'All' || a.applyTo === 'Formats' ? `Borrar ${a.applyTo === 'All' ? 'contenido y formato' : 'el formato'} de ${a.address}` : null),
    label: (a) => `Borrando ${a.address}`,
    async run(args) {
      const applyTo = args.applyTo || 'Contents';
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        let address;
        if (applyTo === 'Formats') {
          range.load('address');
          await ctx.sync();
          address = range.address;
          recordNonUndoable('Formato borrado', address);
        } else {
          address = await snapshotRange(ctx, range, applyTo === 'All' ? 'Contenido y formato borrados' : 'Contenido borrado');
        }
        range.clear(applyTo);
        await ctx.sync();
        return { address, cleared: applyTo };
      });
    },
  },

  /* ---------- Formato ---------- */
  {
    name: 'format_range',
    description: 'Aplica formato a un rango: negrita, cursiva, colores, tamaño, alineación, formato numérico, ajuste de texto y bordes.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        format: {
          type: 'object',
          properties: {
            bold: { type: 'boolean' },
            italic: { type: 'boolean' },
            fontColor: { type: 'string', description: 'Hex "#FF0000" o nombre' },
            fillColor: { type: 'string', description: 'Color de relleno (hex o nombre)' },
            fontSize: { type: 'number' },
            horizontalAlignment: { type: 'string', enum: ['General', 'Left', 'Center', 'Right'] },
            numberFormat: { type: 'string', description: 'Ej: "#,##0.00", "0.0%", "dd/mm/yyyy", "$#,##0"' },
            wrapText: { type: 'boolean' },
            borders: { type: 'string', enum: ['None', 'Outline', 'All'], description: 'Bordes finos' },
            borderColor: { type: 'string' },
          },
        },
      },
      required: ['address', 'format'],
    },
    label: (a) => `Dando formato a ${a.address}`,
    async run(args) {
      const f = args.format || {};
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        range.load('address, rowCount, columnCount');
        await ctx.sync();
        const fmt = range.format;
        if (f.bold !== undefined) fmt.font.bold = f.bold;
        if (f.italic !== undefined) fmt.font.italic = f.italic;
        if (f.fontColor || f.color) fmt.font.color = f.fontColor || f.color;
        if (f.fillColor) fmt.fill.color = f.fillColor;
        if (f.fontSize) fmt.font.size = f.fontSize;
        if (f.horizontalAlignment) fmt.horizontalAlignment = f.horizontalAlignment;
        if (f.wrapText !== undefined) fmt.wrapText = f.wrapText;
        if (f.numberFormat) {
          range.numberFormat = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(f.numberFormat));
        }
        if (f.borders || f.borderColor) {
          const edges = ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'];
          if (f.borders === 'All') edges.push('InsideHorizontal', 'InsideVertical');
          for (const e of edges) {
            const b = fmt.borders.getItem(e);
            b.style = f.borders === 'None' ? 'None' : 'Continuous';
            if (f.borderColor && f.borders !== 'None') b.color = f.borderColor;
          }
        }
        await ctx.sync();
        recordNonUndoable('Formato aplicado', range.address);
        return { address: range.address, formatted: true };
      });
    },
  },
  {
    name: 'set_column_width',
    description: 'Ajusta el ancho (en puntos) de columnas, ej: "A:C".',
    parameters: { type: 'object', properties: { columns: { type: 'string', description: 'Ej: "B:B" o "Ventas!A:D"' }, width: { type: 'number' }, sheetName: SHEET_PARAM }, required: ['columns', 'width'] },
    label: (a) => `Ajustando ancho de ${a.columns}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        getRange(ctx, args.columns, args.sheetName).format.columnWidth = args.width;
        await ctx.sync();
        return { columns: args.columns, width: args.width };
      });
    },
  },
  {
    name: 'set_row_height',
    description: 'Ajusta el alto (en puntos) de filas, ej: "1:1".',
    parameters: { type: 'object', properties: { rows: { type: 'string', description: 'Ej: "1:3"' }, height: { type: 'number' }, sheetName: SHEET_PARAM }, required: ['rows', 'height'] },
    label: (a) => `Ajustando alto de ${a.rows}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        getRange(ctx, args.rows, args.sheetName).format.rowHeight = args.height;
        await ctx.sync();
        return { rows: args.rows, height: args.height };
      });
    },
  },
  {
    name: 'auto_fit_columns',
    description: 'Autoajusta el ancho de columnas al contenido (por defecto el rango usado de la hoja).',
    parameters: { type: 'object', properties: { address: { type: 'string', description: 'Rango opcional' }, sheetName: SHEET_PARAM }, required: [] },
    label: () => 'Autoajustando columnas',
    async run(args) {
      return Excel.run(async (ctx) => {
        const rng = args.address ? getRange(ctx, args.address, args.sheetName) : getSheet(ctx, args.sheetName).getUsedRange();
        rng.format.autofitColumns();
        await ctx.sync();
        return { autofitted: true };
      });
    },
  },
  {
    name: 'merge_cells',
    description: 'Combina (o separa) celdas de un rango.',
    parameters: { type: 'object', properties: { address: { type: 'string', description: ADDRESS_DESC }, sheetName: SHEET_PARAM, unmerge: { type: 'boolean' } }, required: ['address'] },
    label: (a) => `${a.unmerge ? 'Separando' : 'Combinando'} ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        if (args.unmerge) range.unmerge(); else range.merge(false);
        await ctx.sync();
        return { address: args.address, merged: !args.unmerge };
      });
    },
  },
  {
    name: 'add_conditional_format',
    description: 'Añade formato condicional: escala de color, barras de datos o regla por valor de celda.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        type: { type: 'string', enum: ['ColorScale', 'DataBar', 'CellValue'] },
        colors: { type: 'array', items: { type: 'string' }, description: 'ColorScale: [mín, medio, máx]. DataBar: [color]. CellValue: [relleno, fuente?]' },
        operator: { type: 'string', enum: ['GreaterThan', 'LessThan', 'EqualTo', 'NotEqualTo', 'Between', 'NotBetween', 'GreaterThanOrEqual', 'LessThanOrEqual'], description: 'Solo CellValue' },
        formula1: { type: 'string', description: 'Solo CellValue: valor o fórmula, ej: "1000" o "=$B$1"' },
        formula2: { type: 'string', description: 'Solo CellValue con Between/NotBetween' },
      },
      required: ['address', 'type'],
    },
    label: (a) => `Formato condicional en ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        const colors = args.colors || [];
        const cf = range.conditionalFormats.add(args.type);
        if (args.type === 'ColorScale') {
          cf.colorScale.criteria = {
            minimum: { formula: null, type: 'LowestValue', color: colors[0] || '#F8696B' },
            midpoint: { formula: '50', type: 'Percentile', color: colors[1] || '#FFEB84' },
            maximum: { formula: null, type: 'HighestValue', color: colors[2] || '#63BE7B' },
          };
        } else if (args.type === 'DataBar') {
          cf.dataBar.positiveFormat.fillColor = colors[0] || '#638EC6';
        } else {
          const rule = { formula1: String(args.formula1 ?? '0'), operator: args.operator || 'GreaterThan' };
          if (args.formula2 !== undefined) rule.formula2 = String(args.formula2);
          cf.cellValue.rule = rule;
          cf.cellValue.format.fill.color = colors[0] || '#FFC7CE';
          cf.cellValue.format.font.color = colors[1] || '#9C0006';
        }
        range.load('address');
        await ctx.sync();
        recordNonUndoable('Formato condicional', range.address);
        return { address: range.address, type: args.type, added: true };
      });
    },
  },

  /* ---------- Tablas, gráficos, orden y filtros ---------- */
  {
    name: 'create_table',
    description: 'Convierte un rango en una tabla de Excel con estilo.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        hasHeaders: { type: 'boolean', description: 'Default: true' },
        name: { type: 'string', description: 'Nombre de la tabla (opcional, sin espacios)' },
        style: { type: 'string', description: 'Ej: "TableStyleMedium2"' },
      },
      required: ['address'],
    },
    label: (a) => `Creando tabla en ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const table = ctx.workbook.tables.add(getRange(ctx, args.address, args.sheetName), args.hasHeaders !== false);
        if (args.name) table.name = args.name;
        if (args.style) table.style = args.style;
        table.load('name');
        const r = table.getRange();
        r.load('address');
        await ctx.sync();
        const name = table.name;
        ChangeTracker.record({
          label: `Tabla ${name} creada`,
          address: r.address,
          undo: () => Excel.run(async (c) => { c.workbook.tables.getItem(name).convertToRange(); await c.sync(); }),
        });
        return { name, address: r.address };
      });
    },
  },
  {
    name: 'create_chart',
    description: 'Crea un gráfico a partir de un rango (la primera fila/columna se usa como encabezados/categorías).',
    parameters: {
      type: 'object',
      properties: {
        chartType: { type: 'string', enum: ['ColumnClustered', 'ColumnStacked', 'BarClustered', 'Line', 'LineMarkers', 'Pie', 'Doughnut', 'Area', 'XYScatter'] },
        sourceAddress: { type: 'string', description: ADDRESS_DESC },
        sheetName: SHEET_PARAM,
        title: { type: 'string' },
        position: { type: 'string', description: 'Celda donde colocar la esquina superior izquierda, ej: "H2"' },
      },
      required: ['chartType', 'sourceAddress'],
    },
    label: (a) => `Creando gráfico de ${a.sourceAddress}`,
    async run(args) {
      const type = args.chartType === 'Scatter' ? 'XYScatter' : args.chartType;
      return Excel.run(async (ctx) => {
        const source = getRange(ctx, args.sourceAddress, args.sheetName);
        const sheet = source.worksheet;
        sheet.load('name');
        const chart = sheet.charts.add(type, source, 'Auto');
        if (args.title) chart.title.text = args.title;
        if (args.position) chart.setPosition(splitAddress(args.position).ref);
        chart.load('name');
        await ctx.sync();
        const sheetName = sheet.name;
        const name = chart.name;
        ChangeTracker.record({
          label: `Gráfico "${args.title || name}" creado`,
          address: `${quoteSheet(sheetName)}!${splitAddress(args.position || '').ref || splitAddress(args.sourceAddress).ref}`,
          undo: () => Excel.run(async (c) => { c.workbook.worksheets.getItem(sheetName).charts.getItem(name).delete(); await c.sync(); }),
        });
        return { created: true, name, sheet: sheetName, chartType: type };
      });
    },
  },
  {
    name: 'sort_range',
    description: 'Ordena un rango por una columna.',
    parameters: {
      type: 'object',
      properties: {
        address: { type: 'string', description: ADDRESS_DESC + ' Incluye encabezados si los hay.' },
        sheetName: SHEET_PARAM,
        column: { type: 'number', description: 'Número de columna dentro del rango (1 = primera)' },
        ascending: { type: 'boolean', description: 'Default: true' },
        hasHeaders: { type: 'boolean', description: 'Default: true' },
      },
      required: ['address', 'column'],
    },
    label: (a) => `Ordenando ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        const address = await snapshotRange(ctx, range, 'Rango ordenado');
        range.sort.apply([{ key: Math.max(0, (args.column || 1) - 1), ascending: args.ascending !== false }], false, args.hasHeaders !== false);
        await ctx.sync();
        return { address, sorted: true };
      });
    },
  },
  {
    name: 'add_autofilter',
    description: 'Activa el autofiltro en un rango con encabezados.',
    parameters: { type: 'object', properties: { address: { type: 'string', description: ADDRESS_DESC }, sheetName: SHEET_PARAM }, required: ['address'] },
    label: (a) => `Activando filtro en ${a.address}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        range.worksheet.autoFilter.apply(range);
        await ctx.sync();
        return { address: args.address, filterApplied: true };
      });
    },
  },

  /* ---------- Hojas y estructura ---------- */
  {
    name: 'insert_worksheet',
    description: 'Inserta una hoja nueva.',
    parameters: { type: 'object', properties: { name: { type: 'string' }, activate: { type: 'boolean', description: 'Default: true' } }, required: [] },
    label: (a) => `Creando hoja ${a.name || ''}`.trim(),
    async run(args) {
      return Excel.run(async (ctx) => {
        const sheet = args.name ? ctx.workbook.worksheets.add(args.name) : ctx.workbook.worksheets.add();
        if (args.activate !== false) sheet.activate();
        sheet.load('name');
        await ctx.sync();
        const name = sheet.name;
        ChangeTracker.record({
          label: `Hoja "${name}" creada`,
          address: `${quoteSheet(name)}!A1`,
          undo: () => Excel.run(async (c) => { c.workbook.worksheets.getItem(name).delete(); await c.sync(); }),
        });
        return { name };
      });
    },
  },
  {
    name: 'delete_worksheet',
    description: 'Elimina una hoja (requiere confirmación del usuario).',
    parameters: { type: 'object', properties: { name: { type: 'string' } }, required: ['name'] },
    confirm: (a) => `Eliminar la hoja "${a.name}" (no se puede deshacer)`,
    label: (a) => `Eliminando hoja ${a.name}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        ctx.workbook.worksheets.getItem(args.name).delete();
        await ctx.sync();
        recordNonUndoable(`Hoja "${args.name}" eliminada`, null, true);
        return { deleted: args.name };
      });
    },
  },
  {
    name: 'rename_worksheet',
    description: 'Cambia el nombre de una hoja.',
    parameters: { type: 'object', properties: { oldName: { type: 'string' }, newName: { type: 'string' } }, required: ['oldName', 'newName'] },
    label: (a) => `Renombrando ${a.oldName} → ${a.newName}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        ctx.workbook.worksheets.getItem(args.oldName).name = args.newName;
        await ctx.sync();
        const { oldName, newName } = args;
        ChangeTracker.record({
          label: `Hoja renombrada a "${newName}"`,
          address: `${quoteSheet(newName)}!A1`,
          undo: () => Excel.run(async (c) => { c.workbook.worksheets.getItem(newName).name = oldName; await c.sync(); }),
        });
        return { oldName, newName };
      });
    },
  },
  {
    name: 'activate_worksheet',
    description: 'Muestra (activa) una hoja.',
    parameters: { type: 'object', properties: { name: { type: 'string' } }, required: ['name'] },
    label: (a) => `Abriendo hoja ${a.name}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        ctx.workbook.worksheets.getItem(args.name).activate();
        await ctx.sync();
        return { activated: args.name };
      });
    },
  },
  {
    name: 'insert_rows_or_columns',
    description: 'Inserta filas (ej: "5:7") o columnas (ej: "C:D") desplazando el contenido existente.',
    parameters: { type: 'object', properties: { address: { type: 'string', description: 'Filas "5:7" o columnas "C:D" (puede incluir hoja)' }, sheetName: SHEET_PARAM }, required: ['address'] },
    label: (a) => `Insertando ${a.address}`,
    async run(args) {
      const isCols = /^[A-Z]+:[A-Z]+$/i.test(splitAddress(args.address).ref.replace(/\$/g, ''));
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        range.insert(isCols ? 'Right' : 'Down');
        await ctx.sync();
        recordNonUndoable(`${isCols ? 'Columnas' : 'Filas'} insertadas`, args.address, true);
        return { inserted: args.address };
      });
    },
  },
  {
    name: 'delete_rows_or_columns',
    description: 'Elimina filas (ej: "5:7") o columnas (ej: "C:D") (requiere confirmación del usuario).',
    parameters: { type: 'object', properties: { address: { type: 'string', description: 'Filas "5:7" o columnas "C:D" (puede incluir hoja)' }, sheetName: SHEET_PARAM }, required: ['address'] },
    confirm: (a) => `Eliminar ${a.address} (no se puede deshacer desde Ayla)`,
    label: (a) => `Eliminando ${a.address}`,
    async run(args) {
      const isCols = /^[A-Z]+:[A-Z]+$/i.test(splitAddress(args.address).ref.replace(/\$/g, ''));
      return Excel.run(async (ctx) => {
        const range = getRange(ctx, args.address, args.sheetName);
        range.delete(isCols ? 'Left' : 'Up');
        await ctx.sync();
        recordNonUndoable(`${isCols ? 'Columnas' : 'Filas'} eliminadas`, args.address, true);
        return { deleted: args.address };
      });
    },
  },
  {
    name: 'select_range',
    description: 'Selecciona un rango y navega hasta él para mostrárselo al usuario.',
    parameters: { type: 'object', properties: { address: { type: 'string', description: ADDRESS_DESC }, sheetName: SHEET_PARAM }, required: ['address'] },
    label: (a) => `Mostrando ${a.address}`,
    async run(args) {
      return selectAddress(args.address, args.sheetName);
    },
  },
  {
    name: 'add_named_range',
    description: 'Define un nombre para un rango (ej: TasaIVA → Supuestos!B2).',
    parameters: { type: 'object', properties: { name: { type: 'string' }, address: { type: 'string', description: ADDRESS_DESC }, sheetName: SHEET_PARAM }, required: ['name', 'address'] },
    label: (a) => `Definiendo nombre ${a.name}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        ctx.workbook.names.add(args.name, getRange(ctx, args.address, args.sheetName));
        await ctx.sync();
        const { name } = args;
        ChangeTracker.record({
          label: `Nombre "${name}" definido`,
          address: args.address,
          undo: () => Excel.run(async (c) => { c.workbook.names.getItem(name).delete(); await c.sync(); }),
        });
        return { name, address: args.address };
      });
    },
  },
  {
    name: 'protect_sheet',
    description: 'Protege una hoja con contraseña opcional (requiere confirmación del usuario).',
    parameters: { type: 'object', properties: { name: { type: 'string', description: 'Hoja (opcional, la activa)' }, password: { type: 'string' } }, required: [] },
    confirm: (a) => `Proteger la hoja ${a.name || 'activa'}${a.password ? ' con contraseña' : ''}`,
    label: (a) => `Protegiendo hoja ${a.name || 'activa'}`,
    async run(args) {
      return Excel.run(async (ctx) => {
        const sheet = getSheet(ctx, args.name);
        sheet.load('name');
        sheet.protection.protect(null, args.password || undefined);
        await ctx.sync();
        recordNonUndoable(`Hoja "${sheet.name}" protegida`, null);
        return { protected: true, sheet: sheet.name };
      });
    },
  },
];

/* --- API pública --- */

async function selectAddress(address, sheetName) {
  return Excel.run(async (ctx) => {
    const { sheet } = splitAddress(address);
    const range = getRange(ctx, address, sheetName);
    if (sheet || sheetName) getSheet(ctx, sheet || sheetName).activate();
    range.select();
    range.load('address');
    await ctx.sync();
    return { selected: range.address };
  });
}

const TOOL_MAP = Object.fromEntries(TOOLS.map((t) => [t.name, t]));

// Nombres antiguos (v1) → nuevos, por compatibilidad con conversaciones o prompts existentes
const TOOL_ALIASES = {
  go_to_cell: 'select_range',
  get_active_worksheet: 'get_workbook_overview',
  get_workbook_info: 'get_workbook_overview',
  get_worksheet_names: 'get_workbook_overview',
};

const EXCEL_TOOLS_DEFINITIONS = TOOLS.map((t) => ({
  type: 'function',
  function: { name: t.name, description: t.description, parameters: t.parameters },
}));

function getTool(name) {
  return TOOL_MAP[name] || TOOL_MAP[TOOL_ALIASES[name]] || null;
}

function toolLabel(name, args) {
  const t = getTool(name);
  try {
    return t && t.label ? t.label(args || {}) : name;
  } catch {
    return name;
  }
}

function toolConfirmation(name, args) {
  const t = getTool(name);
  return t && t.confirm ? t.confirm(args || {}) : null;
}

async function callExcelTool(name, args) {
  const tool = getTool(name);
  if (!tool) return { error: `Herramienta desconocida: ${name}` };
  try {
    return await tool.run(args || {});
  } catch (err) {
    const code = err && err.code ? ` (${err.code})` : '';
    return { error: `${err && err.message ? err.message : String(err)}${code}`, tool: name };
  }
}

const ExcelTools = Object.fromEntries(TOOLS.map((t) => [t.name, (args) => t.run(args || {})]));
