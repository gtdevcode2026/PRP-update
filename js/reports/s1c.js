// Port of "Slide 12 1st daigram Automation/automation3.py" (id: s1c).
// Reads "TPRM Web-Portal Export", pivots zone_assessing x assessment_status
// (count, margins), ensures ACTIVE/Active/Deprioritized/Duplicate columns
// exist (0-filled if missing), merges case-variant Active columns into
// Active_Total, and writes "PRP_Final_Output3.xlsx" with 3 sheets: Pivot,
// Summary (status overview chart), Active by Zone (per-zone chart).
window.Reports.s1c = async function s1c(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'TPRM Web-Portal Export');
  function cleanKey(k) { return String(k).trim().toLowerCase(); }
  var rows = sheet.rows.map(function (row) {
    var out = {};
    Object.keys(row).forEach(function (k) { out[cleanKey(k)] = row[k]; });
    return out;
  });

  var pivot = E.pivotCount(
    rows,
    function (r) { return String(r.zone_assessing); },
    function (r) { return String(r.assessment_status); },
    { margins: true, marginsName: 'Grand Total' }
  );

  var headers = pivot.headers.slice();
  var dataRows = pivot.rows.map(function (r) { return r.slice(); });

  ['ACTIVE', 'Active', 'Deprioritized', 'Duplicate'].forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      headers.push(col);
      dataRows.forEach(function (r) { r.push(0); });
    }
  });

  var activeUpperIdx = headers.indexOf('ACTIVE');
  var activeIdx = headers.indexOf('Active');
  headers.push('Active_Total');
  dataRows.forEach(function (r) {
    var a1 = activeUpperIdx >= 0 ? (r[activeUpperIdx] || 0) : 0;
    var a2 = activeIdx >= 0 ? (r[activeIdx] || 0) : 0;
    r.push(a1 + a2);
  });

  var pivotHeaderRow = ['Zone'].concat(headers);
  var pivotOutRows = pivot.indexVals.map(function (iv, i) { return [iv].concat(dataRows[i]); });
  var pivotOutGrid = [pivotHeaderRow].concat(pivotOutRows);

  var grandIdx = pivot.indexVals.indexOf('Grand Total');
  var grandRow = dataRows[grandIdx];
  var activeTotalIdx = headers.indexOf('Active_Total');
  var deprioritizedIdx = headers.indexOf('Deprioritized');
  var duplicateIdx = headers.indexOf('Duplicate');
  var summaryGrid = [
    ['Status', 'Count'],
    ['Active', grandRow[activeTotalIdx]],
    ['Deprioritized', grandRow[deprioritizedIdx]],
    ['Duplicate', grandRow[duplicateIdx]],
  ];

  var zoneActiveGrid = [['Zone', 'Active']].concat(
    pivot.indexVals
      .map(function (iv, i) { return [iv, dataRows[i][activeTotalIdx]]; })
      .filter(function (pair) { return pair[0] !== 'Grand Total'; })
  );

  var files = [{
    name: 'PRP_Final_Output3.xlsx',
    sheets: [
      { name: 'Pivot', grid: pivotOutGrid },
      { name: 'Summary', grid: summaryGrid },
      { name: 'Active by Zone', grid: zoneActiveGrid },
    ],
  }];

  // Embedded charts reproduce the original openpyxl native BarChart look
  // (white background, single blue series, value labels above each bar,
  // axis titles) instead of the light in-page preview theme.
  function excelBarChart(categories, values, title, xTitle, yTitle) {
    return {
      traces: [{
        x: categories, y: values, type: 'bar',
        marker: { color: '#4472C4' },
        text: values.map(String), textposition: 'outside', textfont: { color: '#000000', size: 10 },
      }],
      layout: {
        paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
        title: { text: title, font: { color: '#000000', size: 13 } },
        xaxis: { title: xTitle, tickfont: { color: '#000000' }, linecolor: '#808080', showline: true },
        yaxis: { title: yTitle, tickfont: { color: '#000000' }, gridcolor: '#D9D9D9', zerolinecolor: '#808080' },
        showlegend: false,
        margin: { t: 50, r: 20, b: 50, l: 50 },
      },
    };
  }
  var statusRows = summaryGrid.slice(1);
  var zoneRows = zoneActiveGrid.slice(1);
  var chart1 = excelBarChart(
    statusRows.map(function (r) { return r[0]; }), statusRows.map(function (r) { return r[1]; }),
    'Assessment Status Overview', 'Assessment Status', 'Count'
  );
  var chart2 = excelBarChart(
    zoneRows.map(function (r) { return r[0]; }), zoneRows.map(function (r) { return r[1]; }),
    'Active by Zone', 'Zone', 'Count'
  );
  var images = {
    'Summary': await B.renderStyledPng(chart1.traces, chart1.layout, 480, 280),
    'Active by Zone': await B.renderStyledPng(chart2.traces, chart2.layout, 480, 280),
  };

  var workbook = new ExcelJS.Workbook();
  function writeSheet(name, grid) {
    var ws = workbook.addWorksheet(name);
    grid.forEach(function (r) { ws.addRow(r); });
    ws.columns.forEach(function (c) { c.width = 16; });
    return ws;
  }
  writeSheet('Pivot', pivotOutGrid);
  var wsSummary = writeSheet('Summary', summaryGrid);
  var wsZone = writeSheet('Active by Zone', zoneActiveGrid);

  if (images['Summary']) {
    var id1 = workbook.addImage({ base64: images['Summary'], extension: 'png' });
    wsSummary.addImage(id1, { tl: { col: 3, row: 1 }, ext: { width: 480, height: 280 } }); // ~"D2"
  }
  if (images['Active by Zone']) {
    var id2 = workbook.addImage({ base64: images['Active by Zone'], extension: 'png' });
    wsZone.addImage(id2, { tl: { col: 3, row: 1 }, ext: { width: 480, height: 280 } }); // ~"D2"
  }

  var buf = await workbook.xlsx.writeBuffer();

  return {
    ok: true,
    files: [{
      name: 'PRP_Final_Output3.xlsx',
      bytes: buf,
      sheets: [
        { name: 'Pivot', grid: pivotOutGrid },
        { name: 'Summary', grid: summaryGrid },
        { name: 'Active by Zone', grid: zoneActiveGrid },
      ],
    }],
    chartImages: images,
  };
};
