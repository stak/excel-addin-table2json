/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

const CONFIG_SHEET_NAME = '#table2json';

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function alert(msg: string): void {
  Office.context.ui.displayDialogAsync(
    'https://localhost:3000/#' + msg,
    {height: 1, width: Math.max(msg.length / 3, 1)}
  );
}

const transpose = <T>(a: T[][]): T[][] => a[0].map((_, c) => a.map(r => r[c]));

interface TableDetail {
  name: string,
  attrRow: string[],
  isValid: boolean,
  body: Excel.Range
}

async function fetchAllTables(context: Excel.RequestContext): Promise<TableDetail[]> {
  const tables = context.workbook.tables;
  tables.load('items');
  await context.sync();

  const tableAndRanges = tables.items.map(t => {
    const r = t.getRange();
    t.load('name');
    r.load(['rowIndex']);
    
    return [t, r];
  });
  await context.sync();

  const tableAndRanges2 = tableAndRanges.map(([t, r]: [Excel.Table, Excel.Range]) => {
    let a;
    if (r.rowIndex > 0) {
      a = r.getRowsAbove();
      a.load('text');
    } else {
      a = { text: [[]] };
    }
    return [t, r, a];
  });
  await context.sync();

  const tableDetails: TableDetail[] = tableAndRanges2.map(([t, r, a]: [Excel.Table, Excel.Range, Excel.Range]) => ({
    name: t.name,
    attrRow: a.text[0],
    isValid: a.text[0].every(attr => !attr || isValidAttr(attr)) &&
             a.text[0].some(attr => isValidAttr(attr)),
    body: t.getDataBodyRange()
  }));

  tableDetails.forEach(d => d.isValid && d.body.load(['values']));
  await context.sync();

  return tableDetails;
}

async function getBlankSheet(context: Excel.RequestContext, name: string): Promise<Excel.Worksheet> {
  const sheet = context.workbook.worksheets.getItemOrNullObject(CONFIG_SHEET_NAME);
  await context.sync();

  if (sheet.isNullObject) {
    const addedSheet = context.workbook.worksheets.add(CONFIG_SHEET_NAME);
    return addedSheet;
  } else {
    sheet.getUsedRange().getOffsetRange(1, 0).clear('All');
    await context.sync();
    return sheet;
  }
}

function isValidAttr(attrString: string): boolean {
  return /^[a-zA-Z0-9_]+(\.[a-zA-Z0-9_]+)*$/.test(attrString);
}

function buildObject(values: any[][], keys: string[]): any[] {
  const result = [];

  for (let row = 0; row < values.length; ++row) {
    const o = {};
    const cols = values[row];
    
    for (let c = 0; c < cols.length; ++c) {
      const value = cols[c];
      const attr = keys[c];
      if (!attr) {
        continue;
      }

      const attrKeys = attr.split('.');
      let innerObject = o;
      for (let l = 0; l < attrKeys.length - 1; ++l) {
        if (!innerObject[attrKeys[l]]) {
          innerObject[attrKeys[l]] = {};
        }
        innerObject = innerObject[attrKeys[l]];
      }
      innerObject[attrKeys[attrKeys.length - 1]] = value;
    }
    result.push(o);
  }
  return result;
}

function updateSheet(sheet: Excel.Worksheet, tables: TableDetail[], jsonFormatter: (data: any[]) => string[]): void {
  sheet.activate();

  // header
  const headers = ['file', 'table', 'isValid', 'rows', 'json'];
  const headerRange = sheet.getRangeByIndexes(0, 0, headers.length, 1);
  headerRange.values = headers.map(h => [h]);
  headerRange.format.fill.color = '#aaffaa';
  headerRange.getRow(0).format.fill.color = '#aaaaff';

  // metadata
  const metadataRange = sheet.getRangeByIndexes(1, 1, 3, tables.length);
  metadataRange.values = [
    tables.map(t => t.name),
    tables.map(t => t.isValid),
    tables.map(t => t.isValid ? t.body.values.length : 0)
  ];

  // json
  const jsonValues: string[][] = [];
  for (let i = 0; i < tables.length; ++i) {
    if (tables[i].isValid) {
      const tableArray = buildObject(tables[i].body.values, tables[i].attrRow);
      jsonValues.push(jsonFormatter(tableArray));
    } else {
      jsonValues.push([]);
    }
  }
  const rowMaxLength = Math.max(...jsonValues.map(v => v.length));
  const jsonRange = sheet.getRangeByIndexes(4, 1, rowMaxLength, tables.length);
  jsonRange.values = transpose(jsonValues);
}

async function makeJsonSheet(event: Office.AddinCommands.Event): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const tables = await fetchAllTables(context);
      const jsonSheet = await getBlankSheet(context, CONFIG_SHEET_NAME);
      updateSheet(jsonSheet, tables, (data) => {
        return [
          '[',
          ...data.map((o, index) =>
            JSON.stringify(o, null) + (index < data.length - 1 ? ',' : '')),
          ']'
        ];
      });
      context.sync();
    });
  } catch (error) {
    if (typeof error === 'string') {
      alert(error);
    } else if (typeof error.debugInfo) {
      alert(error.debugInfo.toString());
    } else if (typeof error.message === 'string') {
      alert(error.message);
    } else {
      alert('error');
    }
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.makeJsonSheet = makeJsonSheet;
