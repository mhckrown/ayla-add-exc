/* === Ayla Add-Exc — Excel Bridge === */
/* Conjunto de herramientas para que el LLM ejecute operaciones en Excel vía Office.js */

const ExcelTools = {

  /* --- Lectura --- */

  async get_selected_range(args) {
    // args: { includeHeaders: boolean }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load('address, values, formulas, rowCount, columnCount, numberFormat');
      await ctx.sync();
      return {
        address: range.address,
        rows: range.rowCount,
        cols: range.columnCount,
        values: range.values,
        formulas: range.formulas,
        hasHeaders: args?.includeHeaders !== false,
      };
    });
  },

  async get_worksheet_names() {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name, items/position');
      await ctx.sync();
      return sheets.items.map(s => ({ name: s.name, position: s.position }));
    });
  },

  async get_worksheet_data(args) {
    // args: { sheetName?: string, range?: string }
    return Excel.run(async (ctx) => {
      let sheet;
      if (args?.sheetName) {
        sheet = ctx.workbook.worksheets.getItem(args.sheetName);
      } else {
        sheet = ctx.workbook.worksheets.getActiveWorksheet();
      }
      const rng = sheet.getRange(args?.range || 'A1:XFD1048576');
      rng.load('address, values, rowCount, columnCount');
      await ctx.sync();
      // Cap at 10,000 cells to avoid sending too much
      const maxCells = 10000;
      let data = rng.values;
      if (data) {
        const total = rng.rowCount * rng.columnCount;
        if (total > maxCells) {
          const ratio = maxCells / total;
          const newRows = Math.max(1, Math.floor(rng.rowCount * ratio));
          const newCols = Math.max(1, Math.floor(rng.columnCount * ratio));
          const smallRange = sheet.getRange(`${rng.address.split('!')[1] || rng.address}`.split(':')[0] +
            `:${rng.getCell(newRows - 1, newCols - 1).getAddress(true).split('!')[1]}`);
          try {
            smallRange.load('values');
            await ctx.sync();
            data = smallRange.values;
          } catch {}
        }
      }
      return {
        sheet: sheet.name,
        address: rng.address,
        rows: data ? data.length : 0,
        cols: data && data[0] ? data[0].length : 0,
        values: data || [],
        truncated: rng.rowCount * rng.columnCount > maxCells,
      };
    });
  },

  async get_active_worksheet() {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load('name, position');
      await ctx.sync();
      return { name: sheet.name, position: sheet.position };
    });
  },

  /* --- Escritura --- */

  async set_cell_value(args) {
    // args: { address: string, value: string|number, isFormula?: boolean }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.address);
      if (args.isFormula) {
        range.formula = args.value;
      } else {
        range.values = [[args.value]];
      }
      await ctx.sync();
      return { address: args.address, written: true, isFormula: !!args.isFormula };
    });
  },

  async set_range_values(args) {
    // args: { rangeAddress: string, values: any[][] }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      range.values = args.values;
      await ctx.sync();
      return { address: args.rangeAddress, rows: args.values.length, cols: args.values[0]?.length || 0 };
    });
  },

  async add_formula(args) {
    // args: { rangeAddress: string, formula: string }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      range.formula = args.formula;
      await ctx.sync();
      return { address: args.rangeAddress, formula: args.formula };
    });
  },

  async fill_range(args) {
    // args: { rangeAddress: string, direction: 'down'|'right'|'up'|'left' }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      range.autoFill(args.direction || 'down');
      await ctx.sync();
      return { address: args.rangeAddress, filled: true };
    });
  },

  /* --- Formato --- */

  async format_range(args) {
    // args: { rangeAddress: string, format: { bold?, italic?, color?, fillColor?, fontSize?, horizontalAlignment?, numberFormat?, wrapText?, borderColor?, borderStyle? } }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      const fmt = range.format;
      if (args.format.bold !== undefined) fmt.font.bold = args.format.bold;
      if (args.format.italic !== undefined) fmt.font.italic = args.format.italic;
      if (args.format.color) fmt.font.color = args.format.color;
      if (args.format.fillColor) fmt.fill.color = args.format.fillColor;
      if (args.format.fontSize) fmt.font.size = args.format.fontSize;
      if (args.format.horizontalAlignment) fmt.horizontalAlignment = args.format.horizontalAlignment;
      if (args.format.numberFormat) fmt.numberFormat = args.format.numberFormat;
      if (args.format.wrapText !== undefined) fmt.wrapText = args.format.wrapText;
      if (args.format.borderColor) {
        fmt.borders.getItem('EdgeTop').color = args.format.borderColor;
        fmt.borders.getItem('EdgeBottom').color = args.format.borderColor;
        fmt.borders.getItem('EdgeLeft').color = args.format.borderColor;
        fmt.borders.getItem('EdgeRight').color = args.format.borderColor;
      }
      await ctx.sync();
      return { address: args.rangeAddress, formatted: true };
    });
  },

  async set_column_width(args) {
    // args: { columnRange: string, width: number }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.columnRange);
      range.format.columnWidth = args.width;
      await ctx.sync();
      return { column: args.columnRange, width: args.width };
    });
  },

  async set_row_height(args) {
    // args: { rowRange: string, height: number }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rowRange);
      range.format.rowHeight = args.height;
      await ctx.sync();
      return { row: args.rowRange, height: args.height };
    });
  },

  async auto_fit_columns(args) {
    // args: { rangeAddress?: string }
    return Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const rng = args?.rangeAddress ? ws.getRange(args.rangeAddress) : ws.getUsedRange();
      rng.format.autofitColumns();
      await ctx.sync();
      return { autofitted: true };
    });
  },

  async merge_cells(args) {
    // args: { rangeAddress: string }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      range.merge();
      await ctx.sync();
      return { address: args.rangeAddress, merged: true };
    });
  },

  async unmerge_cells(args) {
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      range.unmerge();
      await ctx.sync();
      return { address: args.rangeAddress, unmerged: true };
    });
  },

  /* --- Formato condicional --- */

  async add_conditional_format(args) {
    // args: { rangeAddress: string, type: 'colorScale'|'dataBar'|'cellValue', criteria?: { operator, formula1, formula2? }, colors?: string[] }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      const cf = range.conditionalFormats.add(args.type);
      if (args.type === 'colorScale' || args.type === 'ColorScale') {
        cf.colorScale.criteria = {
          minimum: { type: 'LowestValue', color: args.colors?.[0] || '#F8696B' },
          midpoint: { type: 'Percentile', value: 50, color: args.colors?.[1] || '#FEB254' },
          maximum: { type: 'HighestValue', color: args.colors?.[2] || '#63BE7B' },
        };
      } else if (args.type === 'dataBar' || args.type === 'DataBar') {
        cf.dataBar.fillColor = args.colors?.[0] || '#6C3FC5';
      } else if (args.criteria) {
        const preset = cf.preset;
        preset.criterion = args.criteria.operator || 'GreaterThan';
        preset.formula1 = args.criteria.formula1 || '0';
        if (args.criteria.formula2) preset.formula2 = args.criteria.formula2;
      }
      await ctx.sync();
      return { address: args.rangeAddress, type: args.type, added: true };
    });
  },

  /* --- Tablas --- */

  async create_table(args) {
    // args: { rangeAddress: string, hasHeaders?: boolean, style?: string }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      const table = ctx.workbook.tables.add(range, args.hasHeaders !== false);
      if (args.style) table.style = args.style;
      table.load('name, address');
      await ctx.sync();
      return { name: table.name, address: table.address };
    });
  },

  /* --- Gráficos --- */

  async create_chart(args) {
    // args: { chartType: string, sourceRange: string, targetCell?: string, title?: string }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const chart = sheet.charts.add(args.chartType, sheet.getRange(args.sourceRange), 'Auto');
      if (args.title) chart.title.text = args.title;
      chart.activate();
      await ctx.sync();
      // Optionally move chart to a specific cell
      let position = 'F1';
      if (args.targetCell) {
        position = args.targetCell;
      }
      chart.top = 0;
      chart.left = 0;
      // Get approximate position for target cell
      try {
        const targetRange = sheet.getRange(position);
        targetRange.load('top, left');
        await ctx.sync();
        chart.top = targetRange.top;
        chart.left = targetRange.left;
      } catch {}
      await ctx.sync();
      return { created: true, chartType: args.chartType, source: args.sourceRange, position };
    });
  },

  /* --- Filtros y ordenamiento --- */

  async add_autofilter(args) {
    // args: { rangeAddress: string }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      range.autoFilter.apply();
      await ctx.sync();
      return { address: args.rangeAddress, filterApplied: true };
    });
  },

  async sort_range(args) {
    // args: { rangeAddress: string, column: number, ascending?: boolean }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      const sort = range.sort;
      sort.apply([{ key: args.column, ascending: args.ascending !== false }], false);
      await ctx.sync();
      return { address: args.rangeAddress, sorted: true, column: args.column };
    });
  },

  /* --- Hojas --- */

  async insert_worksheet(args) {
    // args: { name?: string, position?: number }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.add();
      if (args?.name) sheet.name = args.name;
      if (args?.position) sheet.position = args.position;
      sheet.activate();
      await ctx.sync();
      return { name: sheet.name, position: sheet.position };
    });
  },

  async delete_worksheet(args) {
    // args: { name: string }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(args.name);
      sheet.delete();
      await ctx.sync();
      return { deleted: args.name };
    });
  },

  async rename_worksheet(args) {
    // args: { oldName: string, newName: string }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(args.oldName);
      sheet.name = args.newName;
      await ctx.sync();
      return { oldName: args.oldName, newName: args.newName };
    });
  },

  /* --- Filas y columnas --- */

  async insert_rows(args) {
    // args: { rangeAddress: string }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(args.rangeAddress);
      range.insert('Down');
      await ctx.sync();
      return { inserted: true };
    });
  },

  async insert_columns(args) {
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(args.rangeAddress);
      range.insert('Right');
      await ctx.sync();
      return { inserted: true };
    });
  },

  async delete_rows(args) {
    // args: { rangeAddress: string }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(args.rangeAddress);
      range.delete('Up');
      await ctx.sync();
      return { deleted: true };
    });
  },

  /* --- Navegación --- */

  async go_to_cell(args) {
    // args: { address: string }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.address);
      range.select();
      await ctx.sync();
      return { selected: args.address };
    });
  },

  async activate_worksheet(args) {
    // args: { name: string }
    return Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getItem(args.name);
      sheet.activate();
      await ctx.sync();
      return { activated: args.name };
    });
  },

  /* --- Utilidades --- */

  async get_workbook_info() {
    return Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      wb.load('name');
      const sheets = wb.worksheets;
      sheets.load('items/name, items/position');
      await ctx.sync();
      return {
        name: wb.name,
        sheets: sheets.items.map(s => ({ name: s.name, position: s.position })),
        sheetCount: sheets.items.length,
      };
    });
  },

  async protect_sheet(args) {
    // args: { name?: string, password?: string }
    return Excel.run(async (ctx) => {
      const sheet = args?.name
        ? ctx.workbook.worksheets.getItem(args.name)
        : ctx.workbook.worksheets.getActiveWorksheet();
      sheet.protection.protect({ password: args?.password || '' });
      await ctx.sync();
      return { protected: true, sheet: sheet.name };
    });
  },

  async add_named_range(args) {
    // args: { name: string, rangeAddress: string }
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(args.rangeAddress);
      ctx.workbook.names.add(args.name, range);
      await ctx.sync();
      return { name: args.name, address: args.rangeAddress };
    });
  },
};

/* --- Tool definitions in OpenAI function-calling format --- */
const EXCEL_TOOLS_DEFINITIONS = [
  {
    type: 'function',
    function: {
      name: 'get_selected_range',
      description: 'Obtiene los valores, fórmulas y dirección del rango actualmente seleccionado. Útil para entender qué datos está viendo el usuario.',
      parameters: { type: 'object', properties: { includeHeaders: { type: 'boolean', description: 'Si la primera fila tiene encabezados (default: true)' } }, required: [] }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_worksheet_names',
      description: 'Obtiene la lista de todas las hojas del libro activo con sus nombres y posiciones.',
      parameters: { type: 'object', properties: {}, required: [] }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_worksheet_data',
      description: 'Obtiene los datos de una hoja o rango específico del libro.',
      parameters: {
        type: 'object',
        properties: {
          sheetName: { type: 'string', description: 'Nombre de la hoja (opcional, usa hoja activa si no se especifica)' },
          range: { type: 'string', description: 'Rango en formato A1:B10 (opcional, usa toda la hoja si no se especifica)' }
        },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_active_worksheet',
      description: 'Obtiene el nombre de la hoja activa actual.',
      parameters: { type: 'object', properties: {}, required: [] }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_workbook_info',
      description: 'Obtiene información del libro actual: nombre, hojas disponibles.',
      parameters: { type: 'object', properties: {}, required: [] }
    }
  },
  {
    type: 'function',
    function: {
      name: 'set_cell_value',
      description: 'Escribe un valor o fórmula en una celda específica.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Dirección de la celda (ej: "B2")' },
          value: { type: 'string', description: 'Valor o fórmula a escribir' },
          isFormula: { type: 'boolean', description: 'Si true, el valor se interpreta como fórmula (ej: "=SUM(A1:A10)")' }
        },
        required: ['address', 'value']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'set_range_values',
      description: 'Escribe un array de valores 2D en un rango de celdas.',
      parameters: {
        type: 'object',
        properties: {
          rangeAddress: { type: 'string', description: 'Rango en formato A1:C5' },
          values: { type: 'array', items: { type: 'array' }, description: 'Array 2D de valores [[f1c1, f1c2], [f2c1, f2c2]]' }
        },
        required: ['rangeAddress', 'values']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_formula',
      description: 'Inserta una fórmula en una celda o rango.',
      parameters: {
        type: 'object',
        properties: {
          rangeAddress: { type: 'string', description: 'Celda o rango (ej: "C2:C10")' },
          formula: { type: 'string', description: 'Fórmula en español (ej: "=SUMA(A2:A10)")' }
        },
        required: ['rangeAddress', 'formula']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'format_range',
      description: 'Aplica formato a un rango de celdas: negrita, color, relleno, fuente, alineación, número, bordes.',
      parameters: {
        type: 'object',
        properties: {
          rangeAddress: { type: 'string', description: 'Rango a formatear (ej: "A1:C10")' },
          format: {
            type: 'object',
            properties: {
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              color: { type: 'string', description: 'Color de fuente (hex: "#FF0000" o nombre: "Red")' },
              fillColor: { type: 'string', description: 'Color de relleno (hex o nombre)' },
              fontSize: { type: 'number' },
              horizontalAlignment: { type: 'string', enum: ['Left', 'Center', 'Right'] },
              numberFormat: { type: 'string', description: 'Formato de número (ej: "$#,##0.00", "0.00%", "dd/mm/yyyy")' },
              wrapText: { type: 'boolean' },
              borderColor: { type: 'string' }
            }
          }
        },
        required: ['rangeAddress', 'format']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_conditional_format',
      description: 'Aplica formato condicional a un rango: escalas de color, barras de datos o reglas.',
      parameters: {
        type: 'object',
        properties: {
          rangeAddress: { type: 'string', description: 'Rango (ej: "D2:D100")' },
          type: { type: 'string', enum: ['ColorScale', 'DataBar', 'CellValue'], description: 'Tipo de formato condicional' },
          colors: { type: 'array', items: { type: 'string' }, description: 'Colores para escala/barras. 3 colores para ColorScale, 1 para DataBar.' },
          criteria: {
            type: 'object',
            properties: {
              operator: { type: 'string', enum: ['GreaterThan', 'LessThan', 'EqualTo', 'Between', 'GreaterThanOrEqualTo', 'LessThanOrEqualTo', 'NotEqualTo'] },
              formula1: { type: 'string' },
              formula2: { type: 'string' }
            }
          }
        },
        required: ['rangeAddress', 'type']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'create_table',
      description: 'Convierte un rango en tabla formateada de Excel.',
      parameters: {
        type: 'object',
        properties: {
          rangeAddress: { type: 'string', description: 'Rango con datos (ej: "A1:D50")' },
          hasHeaders: { type: 'boolean', description: 'Si la primera fila son encabezados (default: true)' },
          style: { type: 'string', description: 'Estilo de tabla (ej: "TableStyleMedium2")' }
        },
        required: ['rangeAddress']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'create_chart',
      description: 'Crea un gráfico a partir de un rango de datos.',
      parameters: {
        type: 'object',
        properties: {
          chartType: { type: 'string', enum: ['ColumnClustered', 'BarClustered', 'Line', 'Pie', 'Doughnut', 'Area', 'Scatter'], description: 'Tipo de gráfico' },
          sourceRange: { type: 'string', description: 'Rango de datos (ej: "A1:B12")' },
          title: { type: 'string', description: 'Título del gráfico (opcional)' },
          targetCell: { type: 'string', description: 'Celda donde colocar el gráfico (ej: "E2")' }
        },
        required: ['chartType', 'sourceRange']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_autofilter',
      description: 'Activa el auto-filtro en un rango.',
      parameters: {
        type: 'object',
        properties: { rangeAddress: { type: 'string', description: 'Rango (ej: "A1:D50")' } },
        required: ['rangeAddress']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'sort_range',
      description: 'Ordena un rango por una columna.',
      parameters: {
        type: 'object',
        properties: {
          rangeAddress: { type: 'string', description: 'Rango a ordenar (incluyendo encabezados)' },
          column: { type: 'number', description: 'Número de columna (1-based) para ordenar' },
          ascending: { type: 'boolean', description: 'Ascendente (default: true)' }
        },
        required: ['rangeAddress', 'column']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'set_column_width',
      description: 'Ajusta el ancho de una o más columnas.',
      parameters: {
        type: 'object',
        properties: { columnRange: { type: 'string', description: 'Rango de columna (ej: "A:A" o "A:C")' }, width: { type: 'number' } },
        required: ['columnRange', 'width']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'auto_fit_columns',
      description: 'Autoajusta el ancho de columnas al contenido.',
      parameters: {
        type: 'object',
        properties: { rangeAddress: { type: 'string', description: 'Rango opcional (ej: "A1:D50")' } },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'insert_worksheet',
      description: 'Inserta una nueva hoja en el libro.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Nombre para la nueva hoja (opcional)' } },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'rename_worksheet',
      description: 'Cambia el nombre de una hoja.',
      parameters: {
        type: 'object',
        properties: { oldName: { type: 'string' }, newName: { type: 'string' } },
        required: ['oldName', 'newName']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'activate_worksheet',
      description: 'Activa (navega a) una hoja específica.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Nombre de la hoja' } },
        required: ['name']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'merge_cells',
      description: 'Combina celdas en un rango.',
      parameters: {
        type: 'object',
        properties: { rangeAddress: { type: 'string' } },
        required: ['rangeAddress']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'protect_sheet',
      description: 'Protege una hoja con contraseña opcional.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string', description: 'Nombre de hoja (opcional, usa activa)' }, password: { type: 'string' } },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'add_named_range',
      description: 'Define un nombre para un rango de celdas.',
      parameters: {
        type: 'object',
        properties: { name: { type: 'string' }, rangeAddress: { type: 'string' } },
        required: ['name', 'rangeAddress']
      }
    }
  },
];

/* --- Call a tool by name --- */
async function callExcelTool(name, args) {
  const fn = ExcelTools[name];
  if (!fn) {
    throw new Error(`Herramienta desconocida: ${name}`);
  }
  try {
    const result = await fn(args);
    return result;
  } catch (err) {
    return { error: err.message, tool: name, args };
  }
}
