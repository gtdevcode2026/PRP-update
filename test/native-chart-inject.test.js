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
  console.log('NativeChartInject: all assertions passed');
})().catch((e) => { console.error(e); process.exit(1); });
