const assert = require('node:assert');
const ExcelJS = require('../vendor/exceljs.min.js');
const fflate = require('../vendor/fflate.min.js');
globalThis.fflate = fflate;
const NCI = require('../js/native-chart-inject.js');

(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Summary');
  ws.addRow(['Zone', 'Open', 'Overdue']);
  ws.addRow(['Z1', 3, 1]);
  ws.addRow(['Z2', 5, 2]);
  const buf = await wb.xlsx.writeBuffer();

  const def = {
    grouping: 'clustered', legend: true,
    categories: { ref: "'Summary'!$A$2:$A$3", cache: ['Z1', 'Z2'] },
    series: [
      { name: { lit: 'Open' },    values: { ref: "'Summary'!$B$2:$B$3", cache: [3, 5] }, color: '156082' },
      { name: { lit: 'Overdue' }, values: { ref: "'Summary'!$C$2:$C$3", cache: [1, 2] }, color: 'C00000' },
    ],
  };
  const out = NCI.inject(new Uint8Array(buf), [
    { sheetName: 'Summary', def, anchor: { fromCol: 4, fromRow: 0, toCol: 12, toRow: 15 } },
  ]);

  const files = fflate.unzipSync(out);
  const d = (p) => new TextDecoder().decode(files[p]);
  assert.ok(files['xl/charts/chart1.xml'], 'chart part added');
  assert.ok(files['xl/drawings/drawing1.xml'], 'drawing part added');
  assert.ok(files['xl/drawings/_rels/drawing1.xml.rels'], 'drawing rels added');
  assert.ok(d('[Content_Types].xml').includes('/xl/charts/chart1.xml'), 'content-type chart');
  assert.ok(d('[Content_Types].xml').includes('/xl/drawings/drawing1.xml'), 'content-type drawing');
  assert.ok(/<drawing r:id="[^"]+"\/>/.test(d('xl/worksheets/sheet1.xml')), 'sheet has <drawing>');
  assert.ok(files['xl/worksheets/_rels/sheet1.xml.rels'], 'sheet rels added');
  assert.ok(d('xl/worksheets/_rels/sheet1.xml.rels').includes('drawings/drawing1.xml'), 'sheet->drawing rel');
  assert.ok(d('xl/drawings/_rels/drawing1.xml.rels').includes('charts/chart1.xml'), 'drawing->chart rel');

  // --- multi-sheet regression: two charts on two different sheets must get
  //     DISTINCT drawing parts, each linked to its own chart. ---
  const wb2 = new ExcelJS.Workbook();
  const a = wb2.addWorksheet('A'); a.addRow(['k', 'v']); a.addRow(['x', 1]); a.addRow(['y', 2]);
  const b = wb2.addWorksheet('B'); b.addRow(['k', 'v']); b.addRow(['p', 3]); b.addRow(['q', 4]);
  const buf2 = await wb2.xlsx.writeBuffer();
  const mk = (sheet, color) => ({
    grouping: 'clustered', legend: false,
    categories: { ref: "'" + sheet + "'!$A$2:$A$3", cache: ['a', 'b'] },
    series: [{ name: { lit: 'v' }, values: { ref: "'" + sheet + "'!$B$2:$B$3", cache: [1, 2] }, color }],
  });
  const out2 = NCI.inject(new Uint8Array(buf2), [
    { sheetName: 'A', def: mk('A', '111111'), anchor: { fromCol: 3, fromRow: 0, toCol: 10, toRow: 12 } },
    { sheetName: 'B', def: mk('B', '222222'), anchor: { fromCol: 3, fromRow: 0, toCol: 10, toRow: 12 } },
  ]);
  const f2 = fflate.unzipSync(out2);
  const d2 = (p) => new TextDecoder().decode(f2[p]);
  assert.ok(f2['xl/drawings/drawing1.xml'] && f2['xl/drawings/drawing2.xml'], 'two distinct drawing parts');
  assert.ok(f2['xl/charts/chart1.xml'] && f2['xl/charts/chart2.xml'], 'two chart parts');
  // each sheet references a DIFFERENT drawing target
  const relA = d2('xl/worksheets/_rels/sheet1.xml.rels');
  const relB = d2('xl/worksheets/_rels/sheet2.xml.rels');
  const tgtA = (relA.match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/) || [])[1];
  const tgtB = (relB.match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/) || [])[1];
  assert.ok(tgtA && tgtB && tgtA !== tgtB, 'sheets point to distinct drawings (' + tgtA + ' vs ' + tgtB + ')');
  // each drawing links to exactly one chart, and they are different charts
  const chA = (d2('xl/drawings/_rels/' + tgtA + '.rels').match(/charts\/(chart\d+\.xml)/) || [])[1];
  const chB = (d2('xl/drawings/_rels/' + tgtB + '.rels').match(/charts\/(chart\d+\.xml)/) || [])[1];
  assert.ok(chA && chB && chA !== chB, 'drawings link to distinct charts (' + chA + ' vs ' + chB + ')');
  // content types must cover BOTH drawings
  assert.equal((d2('[Content_Types].xml').match(/officedocument\.drawing\+xml/g) || []).length, 2, 'two drawing content-types');
  console.log('NativeChartInject: all assertions passed');
})().catch((e) => { console.error(e); process.exit(1); });
