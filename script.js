(function () {
  const formulaTypeEl = document.getElementById('formula-type');
  const paramsToContainer = document.getElementById('params-to-container');
  const paramsFromContainer = document.getElementById('params-from-container');
  const formulaOutput = document.getElementById('formula-output');
  const btnCopy = document.getElementById('btn-copy');

  const params = {};

  function getParam(key) {
    const el = document.getElementById(key);
    return el ? el.value.trim() : (params[key] || '');
  }

  function setParam(key, value) {
    params[key] = value;
  }

  var refMode = null;

  function ref(sheet, table, column) {
    if (!table || !column) return '';
    return sheet ? sheet + '!' + table + '[' + column + ']' : table + '[' + column + ']';
  }

  function refSheet(sheet, column) {
    if (!column) return '';
    return sheet ? sheet + '!' + column : column;
  }

  function keyValue(col) {
    var s = (col || '').trim();
    if (s.indexOf('[@[') === 0) return s;
    if (!s) return '';
    return '[@[' + s + ']]';
  }

  var xlookupFromBlockCount = 1;
  var xlookupToKeyCount = 1;

  function getToKeyValueString() {
    var parts = [];
    for (var i = 0; i < xlookupToKeyCount; i++) {
      var v = keyValue(getParam('to_key_' + i));
      if (v) parts.push(v);
    }
    return parts.length === 0 ? '' : parts.join('&');
  }

  function getXlookupBlocks() {
    var blocks = [];
    for (var i = 0; i < xlookupFromBlockCount; i++) {
      var sheet = getParam('from_' + i + '_sheet');
      var table = getParam('from_' + i + '_table');
      var keyCol = getParam('from_' + i + '_key_column');
      var returnCol = getParam('from_' + i + '_return_column');
      var lookupArr = refMode === 'sheet' ? refSheet(sheet, keyCol) : ref(sheet, table, keyCol);
      var returnArr = refMode === 'sheet' ? refSheet(sheet, returnCol) : ref(sheet, table, returnCol);
      if (lookupArr && returnArr) blocks.push({ lookup: lookupArr, return: returnArr });
    }
    return blocks;
  }

  function buildNestedXlookup(kv, blocks, ifNotFound) {
    if (blocks.length === 0) return '';
    var inner = ifNotFound || '"Not found"';
    for (var i = blocks.length - 1; i >= 0; i--) {
      inner = 'XLOOKUP(' + kv + ', ' + blocks[i].lookup + ', ' + blocks[i].return + ', ' + inner + ', 0)';
    }
    return '=' + inner;
  }

  function getFormulaString() {
    const type = formulaTypeEl.value;
    switch (type) {
      case 'xlookup': {
        const kv = getToKeyValueString();
        const blocks = getXlookupBlocks();
        const ifNotFound = getParam('from_if_not_found') || '"Not found"';
        if (!kv || blocks.length === 0) return '';
        return buildNestedXlookup(kv, blocks, ifNotFound);
      }
      case 'iferror_xlookup': {
        const kv = getToKeyValueString();
        const blocks = getXlookupBlocks();
        const ifNotFound = getParam('from_if_not_found') || '"Not found"';
        if (!kv || blocks.length === 0) return '';
        const nested = buildNestedXlookup(kv, blocks, ifNotFound);
        if (!nested) return '';
        return '=IFERROR(' + nested.slice(1) + ', ' + ifNotFound + ')';
      }
      case 'filter': {
        const fromSheet = getParam('from_sheet');
        const fromTable = getParam('from_table');
        const fromRange = getParam('from_range');
        var array = '';
        if (refMode === 'table' && fromTable) array = fromSheet ? fromSheet + '!' + fromTable : fromTable;
        else if (refMode === 'sheet') array = fromRange ? (fromSheet ? fromSheet + '!' + fromRange : fromRange) : '';
        else array = fromTable ? (fromSheet ? fromSheet + '!' + fromTable : fromTable) : (fromRange || '');
        const include = getParam('from_filter_rule');
        const if_empty = getParam('from_if_empty') || '"None"';
        if (!array || !include) return '';
        return '=FILTER(' + array + ', ' + include + ', ' + if_empty + ')';
      }
      case 'vstack_filter': {
        const s1 = getParam('from1_sheet');
        const t1 = getParam('from1_table');
        const r1 = getParam('from1_range');
        var arr1 = refMode === 'table' && t1 ? (s1 ? s1 + '!' + t1 : t1) : (r1 ? (s1 ? s1 + '!' + r1 : r1) : '');
        const inc1 = getParam('from1_rule');
        const emp1 = getParam('from1_if_empty') || '"None"';
        const s2 = getParam('from2_sheet');
        const t2 = getParam('from2_table');
        const r2 = getParam('from2_range');
        var arr2 = refMode === 'table' && t2 ? (s2 ? s2 + '!' + t2 : t2) : (r2 ? (s2 ? s2 + '!' + r2 : r2) : '');
        const inc2 = getParam('from2_rule');
        const emp2 = getParam('from2_if_empty') || '"None"';
        if (!arr1 || !inc1 || !arr2 || !inc2) return '';
        return '=VSTACK(FILTER(' + arr1 + ', ' + inc1 + ', ' + emp1 + '), FILTER(' + arr2 + ', ' + inc2 + ', ' + emp2 + '))';
      }
      case 'if': {
        const condition = getParam('to_condition');
        const value_if_true = getParam('from_value_true');
        const value_if_false = getParam('from_value_false');
        if (!condition || value_if_true === '' || value_if_false === '') return '';
        return '=IF(' + condition + ', ' + value_if_true + ', ' + value_if_false + ')';
      }
      default:
        return '';
    }
  }

  function addField(container, id, labelText, placeholder) {
    const label = document.createElement('label');
    label.className = 'label';
    label.htmlFor = id;
    label.textContent = labelText;
    const input = document.createElement('input');
    input.type = 'text';
    input.id = id;
    input.className = 'input';
    input.placeholder = placeholder;
    if (params[id] !== undefined) input.value = params[id];
    input.addEventListener('input', function () {
      setParam(id, input.value);
      updateOutput();
    });
    container.appendChild(label);
    container.appendChild(input);
  }

  function addBlockLabel(container, text) {
    const p = document.createElement('p');
    p.className = 'params-block-label';
    p.textContent = text;
    container.appendChild(p);
  }

  function addHintTo(container, text) {
    const p = document.createElement('p');
    p.className = 'hint params-hint';
    p.textContent = text;
    container.appendChild(p);
  }

  function addButton(container, text, className, onClick) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = className || 'btn-add-sheet';
    btn.textContent = text;
    btn.addEventListener('click', onClick);
    container.appendChild(btn);
  }

  function addCardAction(container, text, className, onClick) {
    const wrap = document.createElement('div');
    wrap.className = 'params-card-actions' + (container === paramsToContainer ? ' params-card-actions-left' : '');
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = className || 'btn-add-sheet';
    btn.textContent = text;
    btn.addEventListener('click', onClick);
    wrap.appendChild(btn);
    container.appendChild(wrap);
  }

  function setHints(toText, fromText) {
    var toHint = document.getElementById('params-to-hint');
    var fromHint = document.getElementById('params-from-hint');
    if (toHint) toHint.textContent = toText || '';
    if (fromHint) fromHint.textContent = fromText || '';
  }

  function renderFields() {
    if (!refMode) return;
    paramsToContainer.innerHTML = '';
    paramsFromContainer.innerHTML = '';
    const type = formulaTypeEl.value;
    const isSheet = refMode === 'sheet';
    const isTable = refMode === 'table';

    if (type === 'xlookup' || type === 'iferror_xlookup') {
      setHints('Header or column name', isTable ? 'Sheet name or table name, column names' : 'Sheet name, column ranges');
      for (var j = 0; j < xlookupToKeyCount; j++) {
        addField(paramsToContainer, 'to_key_' + j, 'Header With Key Value', '');
        if (j > 0) {
          addButton(paramsToContainer, 'REMOVE', 'btn-remove-key', (function (idx) {
            return function () {
              var n = xlookupToKeyCount;
              for (var k = 0; k < n; k++) params['to_key_' + k] = getParam('to_key_' + k);
              for (k = idx; k < n - 1; k++) params['to_key_' + k] = params['to_key_' + (k + 1)];
              delete params['to_key_' + (n - 1)];
              xlookupToKeyCount--;
              renderFields();
            };
          })(j));
        }
      }
      addCardAction(paramsToContainer, 'ADD KEY', 'btn-add-sheet', function () {
        xlookupToKeyCount++;
        renderFields();
      });
      for (var i = 0; i < xlookupFromBlockCount; i++) {
        if (i > 0) addBlockLabel(paramsFromContainer, 'Source ' + (i + 1));
        addField(paramsFromContainer, 'from_' + i + '_key_column', 'Match Header With Key Value', '');
        if (isTable) addField(paramsFromContainer, 'from_' + i + '_table', 'Table Name', '');
        else addField(paramsFromContainer, 'from_' + i + '_sheet', 'Sheet name', '');
        addField(paramsFromContainer, 'from_' + i + '_return_column', 'Fetch Header Value', '');
        if (i > 0) {
          addButton(paramsFromContainer, 'REMOVE', 'btn-remove-sheet', (function (idx) {
            return function () {
              var n = xlookupFromBlockCount;
              for (var k = 0; k < n; k++) {
                params['from_' + k + '_sheet'] = getParam('from_' + k + '_sheet');
                params['from_' + k + '_table'] = getParam('from_' + k + '_table');
                params['from_' + k + '_key_column'] = getParam('from_' + k + '_key_column');
                params['from_' + k + '_return_column'] = getParam('from_' + k + '_return_column');
              }
              for (k = idx; k < n - 1; k++) {
                params['from_' + k + '_sheet'] = params['from_' + (k + 1) + '_sheet'];
                params['from_' + k + '_table'] = params['from_' + (k + 1) + '_table'];
                params['from_' + k + '_key_column'] = params['from_' + (k + 1) + '_key_column'];
                params['from_' + k + '_return_column'] = params['from_' + (k + 1) + '_return_column'];
              }
              delete params['from_' + (n - 1) + '_sheet'];
              delete params['from_' + (n - 1) + '_table'];
              delete params['from_' + (n - 1) + '_key_column'];
              delete params['from_' + (n - 1) + '_return_column'];
              xlookupFromBlockCount--;
              renderFields();
            };
          })(i));
        }
      }
      addField(paramsFromContainer, 'from_if_not_found', 'If Not Found', '');
      addCardAction(paramsFromContainer, 'ADD SHEET', 'btn-add-sheet', function () {
        xlookupFromBlockCount++;
        renderFields();
      });
    } else if (type === 'filter') {
      setHints('', isTable ? 'Table name and rule' : 'Sheet name and rule');
      addBlockLabel(paramsFromContainer, 'Source');
      if (isTable) {
        addField(paramsFromContainer, 'from_table', 'Table Name', '');
        addField(paramsFromContainer, 'from_filter_rule', 'Rule', '');
      } else {
        addField(paramsFromContainer, 'from_sheet', 'Sheet name', '');
        addField(paramsFromContainer, 'from_range', 'Range', '');
        addField(paramsFromContainer, 'from_filter_rule', 'Rule', '');
      }
      addField(paramsFromContainer, 'from_if_empty', 'If empty', '');
    } else if (type === 'vstack_filter') {
      setHints('', isTable ? 'Two tables and rules' : 'Two ranges and rules');
      addBlockLabel(paramsFromContainer, 'Block 1');
      if (isTable) {
        addField(paramsFromContainer, 'from1_table', 'Table Name', '');
        addField(paramsFromContainer, 'from1_rule', 'Rule', '');
      } else {
        addField(paramsFromContainer, 'from1_sheet', 'Sheet name', '');
        addField(paramsFromContainer, 'from1_range', 'Range', '');
        addField(paramsFromContainer, 'from1_rule', 'Rule', '');
      }
      addField(paramsFromContainer, 'from1_if_empty', 'If empty', '');
      addBlockLabel(paramsFromContainer, 'Block 2');
      if (isTable) {
        addField(paramsFromContainer, 'from2_table', 'Table Name', '');
        addField(paramsFromContainer, 'from2_rule', 'Rule', '');
      } else {
        addField(paramsFromContainer, 'from2_sheet', 'Sheet name', '');
        addField(paramsFromContainer, 'from2_range', 'Range', '');
        addField(paramsFromContainer, 'from2_rule', 'Rule', '');
      }
      addField(paramsFromContainer, 'from2_if_empty', 'If empty', '');
    } else if (type === 'if') {
      setHints('Condition', 'Then value, else value');
      addField(paramsToContainer, 'to_condition', 'Condition', '');
      addField(paramsFromContainer, 'from_value_true', 'Then', '');
      addField(paramsFromContainer, 'from_value_false', 'Else', '');
    }

    updateOutput();
  }

  function updateOutput() {
    const formula = getFormulaString();
    formulaOutput.textContent = formula || '—';
    btnCopy.disabled = !formula;
  }

  function copyToClipboard() {
    const formula = getFormulaString();
    if (!formula) return;
    navigator.clipboard.writeText(formula).then(
      function () {
        btnCopy.textContent = 'Copied';
        btnCopy.classList.add('copied');
        btnCopy.setAttribute('aria-label', 'Copied to clipboard');
        setTimeout(function () {
          btnCopy.textContent = 'Copy';
          btnCopy.classList.remove('copied');
          btnCopy.setAttribute('aria-label', 'Copy formula to clipboard');
        }, 2000);
      },
      function () {
        btnCopy.classList.remove('copied');
      }
    );
  }

  /* Custom formula-type dropdown */
  const trigger = document.getElementById('formula-type-trigger');
  const list = document.getElementById('formula-type-list');
  const dropdownValue = trigger && trigger.querySelector('.dropdown-value');
  const options = list && list.querySelectorAll('[role="option"]');

  function syncTriggerText() {
    if (!dropdownValue || !formulaTypeEl) return;
    const opt = formulaTypeEl.options[formulaTypeEl.selectedIndex];
    dropdownValue.textContent = opt ? opt.textContent : '';
  }

  function closeList() {
    if (!list || !trigger) return;
    list.hidden = true;
    trigger.setAttribute('aria-expanded', 'false');
    if (options) {
      options.forEach(function (li) {
        li.setAttribute('aria-selected', li.getAttribute('data-value') === formulaTypeEl.value ? 'true' : 'false');
      });
    }
  }

  function openList() {
    if (!list || !trigger) return;
    list.hidden = false;
    trigger.setAttribute('aria-expanded', 'true');
    if (options) {
      options.forEach(function (li) {
        li.setAttribute('aria-selected', li.getAttribute('data-value') === formulaTypeEl.value ? 'true' : 'false');
      });
    }
  }

  var paramsRow = document.getElementById('params-row');
  var outputCard = document.getElementById('output-card');
  var btnModeSheet = document.getElementById('mode-sheet');
  var btnModeTable = document.getElementById('mode-table');

  function setMode(mode) {
    refMode = mode;
    if (paramsRow) paramsRow.classList.remove('params-row-hidden');
    if (outputCard) outputCard.classList.remove('card-output-hidden');
    if (btnModeSheet) btnModeSheet.setAttribute('aria-pressed', mode === 'sheet' ? 'true' : 'false');
    if (btnModeTable) btnModeTable.setAttribute('aria-pressed', mode === 'table' ? 'true' : 'false');
    renderFields();
  }

  if (btnModeSheet) btnModeSheet.addEventListener('click', function () { setMode('sheet'); });
  if (btnModeTable) btnModeTable.addEventListener('click', function () { setMode('table'); });

  formulaTypeEl.addEventListener('change', function () {
    syncTriggerText();
    if (refMode) renderFields();
  });

  if (trigger && list) {
    syncTriggerText();
    trigger.addEventListener('click', function () {
      if (list.hidden) openList(); else closeList();
    });
    if (options) {
      options.forEach(function (li) {
        li.addEventListener('click', function () {
          const val = li.getAttribute('data-value');
          if (val && formulaTypeEl.value !== val) {
            formulaTypeEl.value = val;
            formulaTypeEl.dispatchEvent(new Event('change'));
          }
          syncTriggerText();
          closeList();
        });
      });
    }
    document.addEventListener('click', function (e) {
      if (list.hidden) return;
      const dropdown = document.getElementById('formula-type-dropdown');
      if (dropdown && !dropdown.contains(e.target)) closeList();
    });
    list.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        closeList();
        trigger.focus();
      }
    });
  }

  btnCopy.addEventListener('click', copyToClipboard);
})();
