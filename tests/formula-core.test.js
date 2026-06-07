const test = require('node:test');
const assert = require('node:assert/strict');
const core = require('../formula-core.js');

test('quotes sheet names only when Excel requires it', function () {
  assert.equal(core.sheetRef('Sales', 'A:A'), 'Sales!A:A');
  assert.equal(core.sheetRef('Sales Data', 'A:A'), "'Sales Data'!A:A");
  assert.equal(core.sheetRef("Manager's Data", 'A:A'), "'Manager''s Data'!A:A");
});

test('normalizes table and current-row column references', function () {
  assert.equal(core.tableColumnRef('tblSales', 'Status'), 'tblSales[Status]');
  assert.equal(core.tableColumnRef('tblSales', '[Status]'), 'tblSales[Status]');
  assert.equal(core.currentRowRef('Customer ID'), '[@[Customer ID]]');
  assert.equal(core.currentRowRef('[@[Customer ID]]'), '[@[Customer ID]]');
});

test('formats friendly fallback values as valid Excel arguments', function () {
  assert.equal(core.excelValue('Not found'), '"Not found"');
  assert.equal(core.excelValue('=NA()'), 'NA()');
  assert.equal(core.excelValue('0'), '0');
  assert.equal(core.excelValue('A2'), 'A2');
  assert.equal(core.excelValue('He said "no"'), '"He said ""no"""');
});

test('qualifies shorthand table filter rules', function () {
  assert.equal(
    core.qualifyTableRule('[Status]="Active"', 'tblSales'),
    'tblSales[Status]="Active"'
  );
  assert.equal(
    core.qualifyTableRule('tblSales[Status]="Active"', 'tblSales'),
    'tblSales[Status]="Active"'
  );
});

test('builds composite and fallback XLOOKUP chains', function () {
  const sources = [
    { lookup: 'tblOne[ID]&tblOne[Date]', returnValue: 'tblOne[Value]' },
    { lookup: 'tblTwo[ID]&tblTwo[Date]', returnValue: 'tblTwo[Value]' }
  ];

  assert.equal(
    core.nestedXlookup('[@[ID]]&[@[Date]]', sources, '"Missing"'),
    'XLOOKUP([@[ID]]&[@[Date]], tblOne[ID]&tblOne[Date], tblOne[Value], XLOOKUP([@[ID]]&[@[Date]], tblTwo[ID]&tblTwo[Date], tblTwo[Value], "Missing"))'
  );
});
