const assert = require('node:assert');
const NC = require('../js/native-chart.js');

const def = {
  grouping: 'clustered', legend: true,
  categories: { ref: "'S'!$A$2:$A$3", cache: ['Z1', 'Z2'] },
  series: [
    { name: { ref: "'S'!$B$1", lit: 'Open' },  values: { ref: "'S'!$B$2:$B$3", cache: [3, 5] }, color: '156082' },
    { name: { ref: "'S'!$C$1", lit: 'Overdue' }, values: { ref: "'S'!$C$2:$C$3", cache: [1, 2] }, color: 'C00000' },
  ],
};
const xml = NC.chartXml(def);
assert.ok(xml.includes('<c:barChart>'), 'has barChart');
assert.ok(xml.includes('<c:grouping val="clustered"/>'), 'clustered');
assert.ok(xml.includes('<a:srgbClr val="156082"/>'), 'series1 color');
assert.ok(xml.includes('<a:srgbClr val="C00000"/>'), 'series2 color');
assert.ok(xml.includes('<c:v>Open</c:v>'), 'series1 name cache');
assert.ok(xml.includes("<c:f>'S'!$B$2:$B$3</c:f>"), 'series1 value ref');
assert.ok(xml.includes('<c:legend>'), 'legend present');
assert.match(xml, /^<\?xml/, 'xml decl');

// stacked
const st = NC.chartXml(Object.assign({}, def, { grouping: 'stacked' }));
assert.ok(st.includes('<c:grouping val="stacked"/>'), 'stacked grouping');
assert.ok(st.includes('<c:overlap val="100"/>'), 'stacked overlap 100');

// no legend
const nl = NC.chartXml(Object.assign({}, def, { legend: false }));
assert.ok(!nl.includes('<c:legend>'), 'no legend');

// per-point colors
const pp = NC.chartXml({ grouping: 'clustered', legend: false,
  categories: { ref: "'S'!$A$2:$A$4", cache: ['a','b','c'] },
  series: [ { name: { lit: 'v' }, values: { ref: "'S'!$B$2:$B$4", cache: [1,2,3] }, color: 'FFC000',
    points: [ {idx:0, color:'BFBFBF'}, {idx:2, color:'BFBFBF'} ] } ] });
assert.ok(pp.includes('<c:dPt>'), 'has data points');
assert.ok(pp.includes('<a:srgbClr val="BFBFBF"/>'), 'per-point color');

// horizontal bar + percent axis + reversed categories
const hb = NC.chartXml({ grouping: 'clustered', legend: false, barDir: 'bar',
  catReversed: true, valNumFmt: '0%',
  dataLabels: { position: 'outEnd', numFmt: '0%', color: '000000' },
  categories: { ref: "'S'!$A$1:$A$2", cache: ['m1', 'm2'] },
  series: [ { name: { lit: 'kpi' }, values: { ref: "'S'!$B$1:$B$2", cache: [0.5, 0.8] }, color: '4472C4' } ] });
assert.ok(hb.includes('<c:barDir val="bar"/>'), 'horizontal barDir');
assert.ok(hb.includes('<c:numFmt formatCode="0%" sourceLinked="0"/>'), 'percent numfmt');
assert.ok(hb.includes('<c:orientation val="maxMin"/>'), 'reversed category orientation');
assert.ok(/<c:catAx>[\s\S]*<c:axPos val="l"\/>/.test(hb), 'cat axis on left for bar');
assert.ok(/<c:valAx>[\s\S]*<c:axPos val="b"\/>/.test(hb), 'val axis on bottom for bar');

// plot hidden cells (helper blocks rely on this)
assert.ok(xml.includes('<c:plotVisOnly val="0"/>'), 'plotVisOnly=0');

// drawing
const d = NC.drawingXml({ fromCol: 6, fromRow: 1, toCol: 14, toRow: 21 }, 'rId1');
assert.ok(d.includes('<xdr:twoCellAnchor>'), 'anchor');
assert.ok(d.includes('<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>'), 'chart rel');
assert.ok(d.includes('<xdr:col>6</xdr:col>'), 'from col');
console.log('NativeChart: all assertions passed');
