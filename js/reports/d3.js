// Port of "daigram 3 automation/automation.py" (id: d3; the embedded/shipped
// version uses relative paths — "PRP Sample Jun (2).xlsx" / "output.xlsx" —
// the standalone .py on disk has stale absolute paths from the original
// author's machine, not used here).
// Reads "OneTrust Assessment", auto-detects Stage/Organization/Working1/
// Working2 columns by substring, filters Working2=="Beyond 1 Year Overdue" &
// Working1 in [Pending, Completed in 2026], pivots Organization x Working1
// (count), appends Grand Total column then row, writes "output.xlsx" sheet
// "Summary" with a native stacked-bar-style chart anchored ~H5.
window.Reports.d3 = async function d3(wb) {
  var E = window.ReportEngine;
  var B = window.ReportBridge;

  var sheet = E.readSheet(wb, 'OneTrust Assessment');
  var rawHeaders = sheet.headers;
  var trimmedHeaders = rawHeaders.map(function (h) { return String(h).trim(); });
  var rows = sheet.rows.map(function (row) {
    var out = {};
    rawHeaders.forEach(function (h, i) { out[trimmedHeaders[i]] = row[h]; });
    return out;
  });

  function findCol(substr) {
    for (var i = 0; i < trimmedHeaders.length; i++) {
      if (trimmedHeaders[i].indexOf(substr) !== -1) return trimmedHeaders[i];
    }
    throw new Error('Column containing "' + substr + '" not found');
  }
  var orgCol = findCol('Organization');
  var w1Col = findCol('Working1');
  var w2Col = findCol('Working2');
  // stage_col/classify_stage only feed "Final Status", which is used purely
  // as the pivot's *value* column for counting — its actual text (Completed/
  // Open) never appears in the output, so it needn't be computed here.

  var filtered = rows.filter(function (r) {
    return r[w2Col] === 'Beyond 1 Year Overdue' && (r[w1Col] === 'Pending' || r[w1Col] === 'Completed in 2026');
  });

  var pivot = E.pivotCount(
    filtered,
    function (r) { return String(r[orgCol]); },
    function (r) { return String(r[w1Col]); },
    { margins: false }
  );

  // Manual two-step Grand Total (matches `pivot["Grand Total"]=sum(axis=1)`
  // then `pivot.loc["Grand Total"]=sum()` — column first, then row over the
  // now-wider matrix).
  var headers = pivot.headers.concat(['Grand Total']);
  var dataRows = pivot.rows.map(function (r) {
    var sum = r.reduce(function (a, b) { return a + b; }, 0);
    return r.concat([sum]);
  });
  var totalsRow = headers.map(function (_, c) {
    return dataRows.reduce(function (acc, r) { return acc + r[c]; }, 0);
  });
  var indexVals = pivot.indexVals.concat(['Grand Total']);
  dataRows.push(totalsRow);

  var headerRow = [orgCol].concat(headers);
  var gridRows = indexVals.map(function (iv, i) { return [iv].concat(dataRows[i]); });
  var grid = [headerRow].concat(gridRows);

  var files = [{ name: 'output.xlsx', sheets: [{ name: 'Summary', grid: grid }] }];

  // Embedded chart reproduces the original openpyxl native stacked BarChart
  // (white background, "Completed in 2026"=red / "Pending"=orange, no data
  // labels, axis titles) — built from the pivot BEFORE the Grand Total
  // column/row were appended, matching the original's Reference(min_col=2,
  // max_col=3, max_row=ws.max_row-1) which excludes both.
  var D3_COLORS = { 'Completed in 2026': '#FF0000', 'Pending': '#FFC000' };
  var images = {};
  if (pivot.indexVals.length) {
    var d3Traces = pivot.headers.map(function (colName, idx) {
      return {
        x: pivot.indexVals, y: pivot.rows.map(function (r) { return r[idx]; }),
        type: 'bar', name: colName, marker: { color: D3_COLORS[colName] || '#4472C4' },
      };
    });
    var d3Layout = {
      barmode: 'stack',
      paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
      title: { text: 'Assessments Completed vs Open', font: { color: '#000000', size: 13 } },
      xaxis: { title: 'Region', tickfont: { color: '#000000' } },
      yaxis: { title: 'Count', gridcolor: '#D9D9D9' },
      legend: { font: { color: '#000000' } },
      margin: { t: 50, r: 20, b: 60, l: 50 },
    };
    images['Summary'] = await B.renderStyledPng(d3Traces, d3Layout, 480, 280);
  }

  var workbook = new ExcelJS.Workbook();
  var ws = workbook.addWorksheet('Summary');
  grid.forEach(function (r) { ws.addRow(r); });
  ws.columns.forEach(function (c) { c.width = 16; });

  if (images['Summary']) {
    var imgId = workbook.addImage({ base64: images['Summary'], extension: 'png' });
    ws.addImage(imgId, { tl: { col: 7, row: 4 }, ext: { width: 480, height: 280 } }); // ~"H5"
  }

  var buf = await workbook.xlsx.writeBuffer();

  return {
    ok: true,
    files: [{ name: 'output.xlsx', bytes: buf, sheets: [{ name: 'Summary', grid: grid }] }],
    chartImages: images,
  };
};
