(function () {
  const formulaTypeEl = document.getElementById('formula-type');
  const paramsRow = document.getElementById('params-row');
  const paramsToContainer = document.getElementById('params-to-container');
  const paramsFromContainer = document.getElementById('params-from-container');
  const outputCard = document.getElementById('output-card');
  const formulaOutput = document.getElementById('formula-output');
  const formulaStatus = document.getElementById('formula-status');
  const btnCopy = document.getElementById('btn-copy');
  const btnModeSheet = document.getElementById('mode-sheet');
  const btnModeTable = document.getElementById('mode-table');
  const {
    currentRowRef,
    excelValue,
    expression,
    nestedXlookup,
    qualifyTableRule,
    sheetRef,
    tableColumnRef
  } = window.XLFormulaCore;

  const params = {};
  let refMode = null;
  let xlookupSourceCount = 1;
  let xlookupKeyCount = 1;

  function getParam(key) {
    const el = document.getElementById(key);
    return el ? el.value.trim() : (params[key] || '');
  }

  function rememberVisibleParams() {
    document.querySelectorAll('.params-card input').forEach(function (input) {
      params[input.id] = input.value;
    });
  }

  function getLookupValue() {
    const values = [];
    for (let keyIndex = 0; keyIndex < xlookupKeyCount; keyIndex++) {
      const value = getParam('to_key_' + keyIndex);
      if (!value) return '';
      values.push(refMode === 'table' ? currentRowRef(value) : expression(value));
    }
    return values.join('&');
  }

  function getXlookupSources() {
    const sources = [];
    for (let sourceIndex = 0; sourceIndex < xlookupSourceCount; sourceIndex++) {
      const sheet = refMode === 'sheet' ? getParam('from_' + sourceIndex + '_sheet') : '';
      const table = refMode === 'table' ? getParam('from_' + sourceIndex + '_table') : '';
      const lookupParts = [];

      for (let keyIndex = 0; keyIndex < xlookupKeyCount; keyIndex++) {
        const column = getParam('from_' + sourceIndex + '_key_' + keyIndex);
        const lookupRef = refMode === 'sheet'
          ? sheetRef(sheet, column)
          : tableColumnRef(table, column);
        if (!lookupRef) {
          lookupParts.length = 0;
          break;
        }
        lookupParts.push(lookupRef);
      }

      const returnColumn = getParam('from_' + sourceIndex + '_return');
      const returnRef = refMode === 'sheet'
        ? sheetRef(sheet, returnColumn)
        : tableColumnRef(table, returnColumn);

      if (lookupParts.length === xlookupKeyCount && returnRef) {
        sources.push({ lookup: lookupParts.join('&'), returnValue: returnRef });
      }
    }
    return sources;
  }

  function sourceArray(prefix) {
    const sheet = refMode === 'sheet' ? getParam(prefix + '_sheet') : '';
    if (refMode === 'table') return getParam(prefix + '_table');
    return sheetRef(sheet, getParam(prefix + '_range'));
  }

  function filterRule(prefix) {
    const rule = getParam(prefix + '_rule');
    if (refMode === 'table') return qualifyTableRule(rule, getParam(prefix + '_table'));
    const sheet = getParam(prefix + '_sheet');
    const normalized = expression(rule);
    if (!sheet || !normalized) return normalized;
    return normalized.replace(/(?<![A-Za-z0-9_'.!])(\$?[A-Z]{1,3}(?:\$?\d+|:\$?[A-Z]{1,3}(?:\$?\d+)?))/gi, function (range) {
      return sheetRef(sheet, range);
    });
  }

  function getFormula() {
    const type = formulaTypeEl.value;

    if (type === 'xlookup' || type === 'iferror_xlookup') {
      const lookupValue = getLookupValue();
      const sources = getXlookupSources();
      if (!lookupValue || sources.length !== xlookupSourceCount) return '';

      if (type === 'xlookup') return '=' + nestedXlookup(lookupValue, sources, '');

      const fallback = excelValue(getParam('from_if_not_found'), '"Not found"');
      const lookup = nestedXlookup(lookupValue, sources, fallback);
      return '=IFERROR(' + lookup + ', ' + fallback + ')';
    }

    if (type === 'filter') {
      const array = sourceArray('from');
      const include = filterRule('from');
      if (!array || !include) return '';
      return '=FILTER(' + array + ', ' + include + ', ' + excelValue(getParam('from_if_empty'), '"None"') + ')';
    }

    if (type === 'vstack_filter') {
      const array1 = sourceArray('from1');
      const include1 = filterRule('from1');
      const array2 = sourceArray('from2');
      const include2 = filterRule('from2');
      if (!array1 || !include1 || !array2 || !include2) return '';
      return '=VSTACK(' +
        'FILTER(' + array1 + ', ' + include1 + ', ' + excelValue(getParam('from1_if_empty'), '"None"') + '), ' +
        'FILTER(' + array2 + ', ' + include2 + ', ' + excelValue(getParam('from2_if_empty'), '"None"') + ')' +
      ')';
    }

    if (type === 'if') {
      const condition = expression(getParam('to_condition'));
      const valueIfTrue = excelValue(getParam('from_value_true'));
      const valueIfFalse = excelValue(getParam('from_value_false'));
      if (!condition || !valueIfTrue || !valueIfFalse) return '';
      return '=IF(' + condition + ', ' + valueIfTrue + ', ' + valueIfFalse + ')';
    }

    return '';
  }

  function addField(container, id, labelText, placeholder, description) {
    const label = document.createElement('label');
    label.className = 'label';
    label.htmlFor = id;
    label.textContent = labelText;

    const input = document.createElement('input');
    input.type = 'text';
    input.id = id;
    input.className = 'input';
    input.placeholder = placeholder;
    input.autocomplete = 'off';
    if (params[id] !== undefined) input.value = params[id];
    if (description) input.setAttribute('aria-describedby', id + '_help');
    input.addEventListener('input', function () {
      params[id] = input.value;
      updateOutput();
    });

    container.appendChild(label);
    container.appendChild(input);

    if (description) {
      const help = document.createElement('span');
      help.id = id + '_help';
      help.className = 'field-help';
      help.textContent = description;
      container.appendChild(help);
    }
  }

  function addBlockLabel(container, text) {
    const heading = document.createElement('h3');
    heading.className = 'params-block-label';
    heading.textContent = text;
    container.appendChild(heading);
  }

  function addButton(container, text, className, onClick) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = className;
    button.textContent = text;
    button.addEventListener('click', onClick);
    container.appendChild(button);
  }

  function addCardAction(container, text, onClick) {
    const wrap = document.createElement('div');
    wrap.className = 'params-card-actions' + (container === paramsToContainer ? ' params-card-actions-left' : '');
    addButton(wrap, text, 'btn-add-sheet', onClick);
    container.appendChild(wrap);
  }

  function setHints(toText, fromText) {
    document.getElementById('params-to-hint').textContent = toText || '';
    document.getElementById('params-from-hint').textContent = fromText || '';
  }

  function removeLookupKey(index) {
    rememberVisibleParams();
    for (let keyIndex = index; keyIndex < xlookupKeyCount - 1; keyIndex++) {
      params['to_key_' + keyIndex] = params['to_key_' + (keyIndex + 1)] || '';
      for (let sourceIndex = 0; sourceIndex < xlookupSourceCount; sourceIndex++) {
        params['from_' + sourceIndex + '_key_' + keyIndex] =
          params['from_' + sourceIndex + '_key_' + (keyIndex + 1)] || '';
      }
    }
    delete params['to_key_' + (xlookupKeyCount - 1)];
    for (let sourceIndex = 0; sourceIndex < xlookupSourceCount; sourceIndex++) {
      delete params['from_' + sourceIndex + '_key_' + (xlookupKeyCount - 1)];
    }
    xlookupKeyCount--;
    renderFields();
  }

  function removeSource(index) {
    rememberVisibleParams();
    const sourceKeys = ['sheet', 'table', 'return'];
    for (let keyIndex = 0; keyIndex < xlookupKeyCount; keyIndex++) sourceKeys.push('key_' + keyIndex);

    for (let sourceIndex = index; sourceIndex < xlookupSourceCount - 1; sourceIndex++) {
      sourceKeys.forEach(function (key) {
        params['from_' + sourceIndex + '_' + key] = params['from_' + (sourceIndex + 1) + '_' + key] || '';
      });
    }
    sourceKeys.forEach(function (key) {
      delete params['from_' + (xlookupSourceCount - 1) + '_' + key];
    });
    xlookupSourceCount--;
    renderFields();
  }

  function renderXlookupFields(type) {
    const isSheet = refMode === 'sheet';
    setHints(
      isSheet ? 'THE CELL OR VALUE TO FIND' : 'THE CURRENT ROW VALUE TO FIND',
      isSheet ? 'WHERE EXCEL SHOULD SEARCH' : 'THE SOURCE TABLE AND COLUMNS'
    );

    for (let keyIndex = 0; keyIndex < xlookupKeyCount; keyIndex++) {
      addField(
        paramsToContainer,
        'to_key_' + keyIndex,
        xlookupKeyCount > 1 ? 'Lookup value ' + (keyIndex + 1) : 'Lookup value',
        isSheet ? 'e.g. A2' : 'e.g. Customer ID'
      );
      if (keyIndex > 0) {
        addButton(paramsToContainer, 'REMOVE VALUE', 'btn-remove-key', function () {
          removeLookupKey(keyIndex);
        });
      }
    }
    addCardAction(paramsToContainer, 'ADD LOOKUP VALUE', function () {
      rememberVisibleParams();
      xlookupKeyCount++;
      renderFields();
    });

    for (let sourceIndex = 0; sourceIndex < xlookupSourceCount; sourceIndex++) {
      addBlockLabel(paramsFromContainer, 'Source ' + (sourceIndex + 1));
      if (isSheet) {
        addField(paramsFromContainer, 'from_' + sourceIndex + '_sheet', 'Sheet name', 'e.g. Customer Data');
      } else {
        addField(paramsFromContainer, 'from_' + sourceIndex + '_table', 'Table name', 'e.g. tblCustomers');
      }
      for (let keyIndex = 0; keyIndex < xlookupKeyCount; keyIndex++) {
        addField(
          paramsFromContainer,
          'from_' + sourceIndex + '_key_' + keyIndex,
          xlookupKeyCount > 1 ? 'Lookup in ' + (keyIndex + 1) : 'Lookup in',
          isSheet ? 'e.g. A:A or A2:A500' : 'e.g. Customer ID'
        );
      }
      addField(
        paramsFromContainer,
        'from_' + sourceIndex + '_return',
        'Return from',
        isSheet ? 'e.g. D:D or D2:D500' : 'e.g. Customer Name'
      );
      if (sourceIndex > 0) {
        addButton(paramsFromContainer, 'REMOVE SOURCE', 'btn-remove-sheet', function () {
          removeSource(sourceIndex);
        });
      }
    }

    if (type === 'iferror_xlookup') {
      addField(
        paramsFromContainer,
        'from_if_not_found',
        'If not found',
        'Not found',
        'Plain text is quoted automatically. Formulas may start with =.'
      );
    }

    addCardAction(paramsFromContainer, 'ADD SOURCE', function () {
      rememberVisibleParams();
      xlookupSourceCount++;
      renderFields();
    });
  }

  function renderFilterBlock(prefix, label) {
    const isSheet = refMode === 'sheet';
    if (label) addBlockLabel(paramsFromContainer, label);
    if (isSheet) {
      addField(paramsFromContainer, prefix + '_sheet', 'Sheet name', 'e.g. Sales Data');
      addField(paramsFromContainer, prefix + '_range', 'Return range', 'e.g. A2:D500');
      addField(paramsFromContainer, prefix + '_rule', 'Include rule', 'e.g. D2:D500="Active"');
    } else {
      addField(paramsFromContainer, prefix + '_table', 'Table name', 'e.g. tblSales');
      addField(
        paramsFromContainer,
        prefix + '_rule',
        'Include rule',
        'e.g. [Status]="Active"',
        'The table name is added to column references automatically.'
      );
    }
    addField(paramsFromContainer, prefix + '_if_empty', 'If empty', 'None', 'Plain text is quoted automatically.');
  }

  function renderFields() {
    rememberVisibleParams();
    paramsToContainer.innerHTML = '';
    paramsFromContainer.innerHTML = '';

    const type = formulaTypeEl.value;
    const ready = Boolean(refMode && type);
    paramsRow.classList.toggle('params-row-hidden', !ready);
    outputCard.classList.toggle('card-output-hidden', !ready);
    if (!ready) return;

    if (type === 'xlookup' || type === 'iferror_xlookup') {
      renderXlookupFields(type);
    } else if (type === 'filter') {
      setHints('', refMode === 'table' ? 'THE TABLE AND RULE TO APPLY' : 'THE RANGE AND RULE TO APPLY');
      renderFilterBlock('from', 'Source');
    } else if (type === 'vstack_filter') {
      setHints('', refMode === 'table' ? 'TWO TABLES TO FILTER AND COMBINE' : 'TWO RANGES TO FILTER AND COMBINE');
      renderFilterBlock('from1', 'Source 1');
      renderFilterBlock('from2', 'Source 2');
    } else if (type === 'if') {
      setHints('THE TEST EXCEL SHOULD RUN', 'WHAT TO RETURN');
      addField(paramsToContainer, 'to_condition', 'Condition', 'e.g. A2>10 or A2="Yes"');
      addField(paramsFromContainer, 'from_value_true', 'Then', 'Pass', 'Plain text is quoted automatically.');
      addField(paramsFromContainer, 'from_value_false', 'Else', 'Fail', 'Plain text is quoted automatically.');
    }

    updateOutput();
  }

  function updateOutput() {
    const formula = getFormula();
    formulaOutput.textContent = formula || '—';
    formulaOutput.classList.toggle('formula-output-empty', !formula);
    formulaStatus.textContent = formula
      ? 'Ready to paste into Excel.'
      : 'Complete the required fields above to build your formula.';
    formulaStatus.classList.toggle('formula-status-ready', Boolean(formula));
    btnCopy.disabled = !formula;
  }

  async function copyToClipboard() {
    const formula = getFormula();
    if (!formula) return;
    try {
      await navigator.clipboard.writeText(formula);
      btnCopy.textContent = 'Copied';
      btnCopy.classList.add('copied');
      formulaStatus.textContent = 'Copied to clipboard.';
      setTimeout(function () {
        btnCopy.textContent = 'Copy';
        btnCopy.classList.remove('copied');
        formulaStatus.textContent = 'Ready to paste into Excel.';
      }, 1800);
    } catch (_) {
      formulaStatus.textContent = 'Copy failed. Select the formula and copy it manually.';
    }
  }

  function setMode(mode) {
    if (refMode === mode) return;
    rememberVisibleParams();
    refMode = mode;
    btnModeSheet.setAttribute('aria-pressed', mode === 'sheet' ? 'true' : 'false');
    btnModeTable.setAttribute('aria-pressed', mode === 'table' ? 'true' : 'false');
    renderFields();
  }

  btnModeSheet.addEventListener('click', function () { setMode('sheet'); });
  btnModeTable.addEventListener('click', function () { setMode('table'); });
  formulaTypeEl.addEventListener('change', renderFields);
  btnCopy.addEventListener('click', copyToClipboard);
})();
