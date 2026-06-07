(function (root, factory) {
  const core = factory();
  if (typeof module === 'object' && module.exports) module.exports = core;
  root.XLFormulaCore = core;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  function quoteSheetName(sheet) {
    const name = (sheet || '').trim();
    if (!name) return '';
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) return name;
    return "'" + name.replace(/'/g, "''") + "'";
  }

  function sheetRef(sheet, range) {
    if (!range) return '';
    const safeSheet = quoteSheetName(sheet);
    return safeSheet ? safeSheet + '!' + range : range;
  }

  function cleanColumnName(column) {
    let value = (column || '').trim();
    if (!value) return '';
    const structured = value.match(/\[([^\[\]]+)\]\]?$/);
    if (structured) value = structured[1];
    return value;
  }

  function tableColumnRef(table, column) {
    const tableName = (table || '').trim();
    const columnName = cleanColumnName(column);
    if (!tableName || !columnName) return '';
    return tableName + '[' + columnName + ']';
  }

  function currentRowRef(column) {
    const value = (column || '').trim();
    if (!value) return '';
    if (value.indexOf('[@') === 0) return value;
    return '[@[' + cleanColumnName(value) + ']]';
  }

  function expression(raw) {
    const value = (raw || '').trim();
    return value[0] === '=' ? value.slice(1).trim() : value;
  }

  function excelValue(raw, fallback) {
    const value = (raw || '').trim();
    if (!value) return fallback || '';
    const normalized = expression(value);
    if (/^-?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return normalized;
    if (/^(?:TRUE|FALSE|NA\(\)|#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!))$/i.test(normalized)) return normalized;
    if (/^".*"$/.test(normalized)) return normalized;
    if (/^'.*'$/.test(normalized)) {
      return '"' + normalized.slice(1, -1).replace(/"/g, '""') + '"';
    }
    if (/^(?:[A-Za-z_][A-Za-z0-9_.]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(normalized)) return normalized;
    if (/^[A-Za-z_][A-Za-z0-9_.]*\s*\(/.test(normalized)) return normalized;
    if (/[<>=+\-*/&\[\]]/.test(normalized)) return normalized;
    return '"' + normalized.replace(/"/g, '""') + '"';
  }

  function qualifyTableRule(rule, table) {
    const value = expression(rule);
    const tableName = (table || '').trim();
    if (!value || !tableName) return value;
    return value.replace(/(^|[^\w\]])(\[[^\]]+\])/g, function (_, prefix, column) {
      return prefix + tableName + column;
    });
  }

  function nestedXlookup(lookupValue, sources, fallback) {
    if (!lookupValue || sources.length === 0) return '';
    let result = fallback || '';

    for (let index = sources.length - 1; index >= 0; index--) {
      const source = sources[index];
      const notFound = result ? ', ' + result : '';
      result = 'XLOOKUP(' + lookupValue + ', ' + source.lookup + ', ' + source.returnValue + notFound + ')';
    }
    return result;
  }

  return {
    currentRowRef: currentRowRef,
    excelValue: excelValue,
    expression: expression,
    nestedXlookup: nestedXlookup,
    qualifyTableRule: qualifyTableRule,
    quoteSheetName: quoteSheetName,
    sheetRef: sheetRef,
    tableColumnRef: tableColumnRef
  };
});
