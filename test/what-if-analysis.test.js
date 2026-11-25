const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const script = fs.readFileSync(path.join(__dirname, '../src/WhatIfAnalysis.gs'), 'utf8');
const example = {name: 'Example', range: "'Model'!D4:E5", colinput: 'B2'};

function setup({failCalculation = false, lockAvailable = true} = {}) {
  let saved = '[]';
  let scenarioInput = 99;
  let secondInput = 50;
  const outputWrites = [];
  const activatedRanges = [];
  function makeSheet(sheetId) {
    const sheet = {
      getSheetId: () => sheetId,
      getRange(ref, col, rows, cols) {
        if (typeof ref === 'number') return makeRange(ref, col, rows, cols);
        if (ref.startsWith("'Model'!")) return model.getRange(ref.slice(8));
        if (ref.startsWith("'Other'!")) return other.getRange(ref.slice(8));
        switch (ref) {
          case 'B2': return makeRange(2, 2);
          case 'C2': return makeRange(2, 3);
          case 'D4': return makeRange(4, 4);
          case 'D4:E5': return makeRange(4, 4, 2, 2);
          default: throw new Error('Invalid range');
        }
      },
    };
    function makeRange(row, col, rows = 1, cols = 1) {
      return {
        getSheet: () => sheet,
        activate: () => { activatedRanges.push([sheetId, row, col]); },
        getA1Notation: () => 'D4:E5',
        getRow: () => row,
        getColumn: () => col,
        getLastRow: () => row + rows - 1,
        getLastColumn: () => col + cols - 1,
        getNumRows: () => rows,
        getNumColumns: () => cols,
        getFormula: () => '',
        getValue: () => row === 4 ? scenarioInput * secondInput :
          col === 3 ? secondInput : scenarioInput,
        setValue(value) {
          if (col === 3) secondInput = value;
          else scenarioInput = value;
        },
        getValues() {
          if (row === 5 && col === 4) return [[scenarioInput === 198 ? 396 : 7]];
          if (failCalculation) throw new Error('Calculation failed');
          return [[scenarioInput * 2]];
        },
        setValues: (values) => { outputWrites.push(JSON.stringify(values)); },
      };
    }
    return sheet;
  }
  const model = makeSheet(1);
  const other = makeSheet(2);
  const spreadsheet = {
    getActiveSheet: () => model,
    getRangeByName: (name) => ({
      TableRange: other.getRange('D4:E5'),
      InputRate: other.getRange('B2'),
      OverlappingCell: other.getRange('D4'),
    })[name] ?? null,
  };
  const ctx = vm.createContext({
    Logger: {log() {}},
    LockService: {getDocumentLock: () => ({tryLock: () => lockAvailable, releaseLock() {}})},
    SpreadsheetApp: {getActiveSpreadsheet: () => spreadsheet, flush() {}},
    PropertiesService: {
      getDocumentProperties: () => ({
        getProperty: () => saved,
        setProperty: (key, value) => { saved = value; },
      }),
    },
  });
  vm.runInContext(script, ctx);
  return {ctx, saved: () => saved, scenarioInput: () => scenarioInput, outputWrites, activatedRanges};
}

test('upsert saves independently and recalculation restores the input', () => {
  const {ctx, saved, outputWrites, scenarioInput} = setup();
  ctx.upsertDataTable(-1, example);
  assert.deepEqual(JSON.parse(saved()), [example]);
  assert.deepEqual(outputWrites, []);
  ctx.recalcDataTable(0);
  assert.deepEqual(outputWrites, ['[[14]]']);
  assert.equal(scenarioInput(), 99);
});

test('validation rejects either axis inside the table and accepts another sheet', () => {
  const {ctx, saved} = setup();
  assert.throws(() => ctx.upsertDataTable(-1, {...example, colinput: 'D4'}), /Column Input.*outside/);
  assert.throws(() => ctx.validateConfig_({...example, rowinput: 'D4'}), /Row Input.*outside/);
  assert.throws(() => ctx.validateConfig_({...example, rowinput: 'bad'}), /Row Input.*valid cell/);
  assert.doesNotThrow(() => ctx.validateConfig_({...example, colinput: "'Other'!D4"}));
  assert.equal(saved(), '[]');
});

test('import validates the entire replacement before writing', () => {
  const {ctx, saved} = setup();
  ctx.importDataTables(JSON.stringify([example]), false);
  for (const invalidJson of ['{', '{}', '[null]']) {
    assert.throws(() => ctx.importDataTables(invalidJson, false));
    assert.deepEqual(JSON.parse(saved()), [example]);
  }
  const replacement = {...example, name: 'Replacement'};
  assert.throws(() => ctx.importDataTables(JSON.stringify([
    replacement, {...example, rowinput: 'D4'},
  ]), false), /DataTable \(index 2 of 2\): Row Input.*outside/);
  assert.deepEqual(JSON.parse(saved()), [example]);
  ctx.importDataTables(JSON.stringify([replacement]), false);
  assert.deepEqual(JSON.parse(saved()), [replacement]);
});

test('Save & Calculate opens the saved editor after the server succeeds', () => {
  const html = fs.readFileSync(path.join(__dirname, '../src/sidepanel.html'), 'utf8');
  const sidebarJs = html.split('<script>')[1].split('</script>')[0];
  const saveButton = {disabled: false};
  const importButton = {disabled: false};
  const jsonPanel = {hidden: false};
  const addNew = {
    open: true,
    querySelectorAll: () => [saveButton],
    querySelector: (selector) => ({
      textarea: {value: JSON.stringify(example)},
      '.save': saveButton,
      '.json': jsonPanel,
    })[selector],
  };
  let successHandler;
  const recalculated = [];
  const ctx = vm.createContext({
    document: {querySelector: (selector) => ({
      '#addNew': addNew, '#saveAll': importButton,
    })[selector]},
    google: {script: {run: {
      withSuccessHandler(fn) { successHandler = fn; return this; },
      withFailureHandler() { return this; },
      withUserObject() { return this; },
      recalcDataTable(idx) { recalculated.push(idx); },
      upsertDataTable(idx, config) {
        assert.equal(arguments.length, 2);
        assert.equal(idx, -1);
        assert.equal(config.name, example.name);
      },
    }}},
  });
  vm.runInContext(sidebarJs.replace('      init();', ''), ctx);
  vm.runInContext(`
    message = () => {};
    wireEditorUi = () => {};
    renderDataTables = (dt) => { DOMRefs.dt = dt.map(() => ({open: false})); };
    saveDt(-1, true);
  `, ctx);
  assert.equal(saveButton.disabled, true);
  assert.equal(importButton.disabled, true);
  assert.equal(addNew.open, true);
  assert.deepEqual(recalculated, []);
  successHandler([{name: 'Earlier table'}, example]);
  assert.equal(addNew.open, false);
  assert.equal(saveButton.disabled, false);
  assert.equal(importButton.disabled, false);
  assert.equal(jsonPanel.hidden, true);
  assert.equal(vm.runInContext('DOMRefs.dt[1].open', ctx), true);
  assert.deepEqual(recalculated, [1]);
});

test('Recalculate All handles column, row, and bivariate tables', () => {
  const {ctx, outputWrites} = setup();
  ctx.upsertDataTable(-1, example);
  ctx.upsertDataTable(-1, {...example, colinput: null, rowinput: 'B2'});
  ctx.upsertDataTable(-1, {...example, rowinput: 'C2'});
  assert.deepEqual(outputWrites, []);
  ctx.recalcDataTable(Infinity);
  assert.deepEqual(outputWrites, ['[[14]]', '[[396]]', '[[1386]]']);
});

test('failed calculation restores inputs and leaves the saved list intact', () => {
  const {ctx, saved, scenarioInput, outputWrites} = setup({failCalculation: true});
  ctx.upsertDataTable(-1, example);
  assert.throws(
      () => ctx.recalcDataTable(0),
      /Calculation failed/);
  assert.equal(scenarioInput(), 99);
  assert.deepEqual(JSON.parse(saved()), [example]);
  assert.deepEqual(outputWrites, []);
});

test('sidebar import validates JSON, confirms replacement, and keeps failed edits open', async (t) => {
  const html = fs.readFileSync(path.join(__dirname, '../src/sidepanel.html'), 'utf8');
  const sidebarJs = html.split('<script>')[1].split('</script>')[0];
  const backend = setup();
  backend.ctx.upsertDataTable(-1, example);
  const saveButton = {disabled: false};
  const editor = {open: true};
  const jsonInput = {value: '{'};
  let confirmed = false;
  let prompts = 0;
  const dialog = {querySelector: () => ({}), showModal() { prompts++; }};
  let successHandler;
  let failureHandler;
  const ui = vm.createContext({
    document: {querySelector: (selector) => ({
      '#confirmation': dialog,
      '#addNew': {querySelectorAll: () => []},
      '#allDt': jsonInput,
      '#saveAll': saveButton,
      '#advancedDt': editor,
    })[selector]},
    google: {script: {run: {
      withSuccessHandler(fn) { successHandler = fn; return this; },
      withFailureHandler(fn) { failureHandler = fn; return this; },
      importDataTables(json, reorderOnly) {
        assert.equal(typeof reorderOnly, 'boolean', 'Apps Script RPC arguments cannot be functions');
        let dt;
        try {
          dt = backend.ctx.importDataTables(json, reorderOnly);
        } catch (err) {
          failureHandler(err);
          return;
        }
        successHandler(dt);
      },
    }}},
  });
  vm.runInContext(sidebarJs.replace('      init();', ''), ui);
  vm.runInContext('message = () => {}; renderDataTables = () => {};', ui);
  const importClick = () => {
    const previousPrompts = prompts;
    vm.runInContext('importDt()', ui);
    if (prompts !== previousPrompts) {
      dialog.returnValue = confirmed ? 'yes' : 'no';
      vm.runInContext('resolveConfirmation()', ui);
    }
  };
  importClick();
  assert.equal(prompts, 0);
  jsonInput.value = '[]';
  importClick();
  assert.equal(prompts, 1);
  assert.deepEqual(JSON.parse(backend.saved()), [example]);
  confirmed = true;
  jsonInput.value = JSON.stringify([{...example, colinput: 'D4'}]);
  importClick();
  assert.equal(editor.open, true);
  assert.equal(saveButton.disabled, false);
  assert.deepEqual(JSON.parse(backend.saved()), [example]);
  jsonInput.value = '[]';
  importClick();
  assert.equal(editor.open, false);
  assert.equal(backend.saved(), '[]');

  const second = {...example, name: 'Second'};
  backend.ctx.upsertDataTable(-1, example);
  backend.ctx.upsertDataTable(-1, second);
  confirmed = false;
  jsonInput.defaultValue = backend.saved();
  vm.runInContext(`
    DOMRefs.dt = [{}, {}];
    reorderDt(0, 1);
  `, ui);
  assert.equal(prompts, 3);
  assert.deepEqual(JSON.parse(backend.saved()), [second, example]);

  await t.test('Copy saves an adjacent duplicate and opens its editor', {
    todo: 'copyDt passes its callback as the reorderOnly RPC argument',
  }, () => {
    jsonInput.defaultValue = backend.saved();
    vm.runInContext(`
      renderDataTables = (dt) => {
        DOMRefs.dt = dt.map(() => ({querySelector: () => ({focus() {}})}));
      };
      copyDt(0);
    `, ui);
    assert.equal(prompts, 3);
    assert.deepEqual(JSON.parse(backend.saved()), [second, second, example]);
    assert.equal(vm.runInContext('DOMRefs.dt[1].open', ui), true);
  });
});

test('reorder and copy keep every saved key regardless of property order', () => {
  const {ctx, saved, outputWrites} = setup();
  const second = {...example, name: 'Second', metadata: 'saved'};
  const reorderedExample = {colinput: example.colinput, range: example.range, name: example.name};
  [example, second].forEach(config => ctx.upsertDataTable(-1, config));

  ctx.importDataTables(JSON.stringify([second, reorderedExample, second]), true);
  assert.deepEqual(JSON.parse(saved()), [second, example, second]);
  assert.throws(() => ctx.importDataTables(JSON.stringify([second, second]), true), /saved state changed/);
  assert.throws(() => ctx.importDataTables('[]', true), /saved state changed/);
  assert.throws(() => ctx.importDataTables(JSON.stringify([second, example, {...example, rowinput: 'D4'}]), false), /outside/);
  assert.deepEqual(JSON.parse(saved()), [second, example, second]);
  assert.deepEqual(outputWrites, []);
});

test('dragging a handle reorders the row using its saved index', () => {
  const html = fs.readFileSync(path.join(__dirname, '../src/sidepanel.html'), 'utf8');
  const sidebarJs = html.split('<script>')[1].split('</script>')[0];
  const moves = [];
  const ctx = vm.createContext({
    document: {querySelector: () => ({disabled: false})},
  });
  vm.runInContext(sidebarJs.replace('      init();', ''), ctx);
  vm.runInContext(`
    const handle = {};
    const dropClasses = new Set();
    const summary = {
      querySelector: () => handle,
      classList: {
        add: (name) => dropClasses.add(name),
        remove: (...names) => names.forEach(name => dropClasses.delete(name)),
      },
    };
    reorderDt = (...args) => move(...args);
    wireReorderUi(summary, 0);
    const startDrag = handle.ondragstart;
    wireReorderUi(summary, 2);

    const event = {preventDefault() {}, dataTransfer: {setData() {}}};
    startDrag(event);
    summary.ondragover(event);
  `, Object.assign(ctx, {move: (...args) => moves.push(args)}));
  assert.equal(vm.runInContext('dropClasses.has("dropAfter")', ctx), true);
  vm.runInContext('summary.ondrop(event)', ctx);
  assert.deepEqual(moves, [[0, 2]]);
  assert.equal(vm.runInContext('dropClasses.size', ctx), 0);
});

test('named ranges resolve in all fields and still obey overlap validation', () => {
  const {ctx, saved} = setup();
  const shared = {...example, range: 'TableRange', colinput: 'B2', rowinput: 'InputRate'};
  const {dtRange, colInputCell, rowInputCell} = ctx.validateConfig_(shared);
  assert.equal(dtRange.getSheet().getSheetId(), 2);
  assert.equal(colInputCell.getSheet().getSheetId(), 2);
  assert.equal(rowInputCell.getRow(), colInputCell.getRow());
  assert.equal(rowInputCell.getColumn(), colInputCell.getColumn());
  ctx.upsertDataTable(-1, shared);
  ctx.importDataTables(JSON.stringify([shared]));
  assert.deepEqual(JSON.parse(saved()), [shared]);
  assert.throws(
      () => ctx.validateConfig_({...shared, colinput: 'OverlappingCell'}),
      /Column Input.*outside/);
});

test('recalculation declines to write while another session holds the lock', () => {
  const {ctx, outputWrites} = setup({lockAvailable: false});
  ctx.upsertDataTable(-1, example);
  assert.throws(() => ctx.recalcDataTable(0), /running in another sidebar/);
  assert.deepEqual(outputWrites, []);
});

test('the singleton modal starts a concurrent request only after Yes and clears on close', () => {
  const html = fs.readFileSync(path.join(__dirname, '../src/sidepanel.html'), 'utf8');
  const sidebarJs = html.split('<script>')[1].split('</script>')[0];
  const calls = [];
  const pending = [];
  let onSuccess;
  let onFailure;
  const dialog = {querySelector: () => ({}), showModal() {}};
  const ui = vm.createContext({
    document: {querySelector: selector => selector === '#confirmation' ? dialog : null},
    google: {script: {run: {
      withSuccessHandler(fn) { onSuccess = fn; return this; },
      withFailureHandler(fn) { onFailure = fn; return this; },
      recalcDataTable(idx) {
        calls.push(idx);
        pending.push({onSuccess, onFailure});
      },
    }}},
  });
  vm.runInContext(sidebarJs.replace('      init();', ''), ui);
  const run = code => vm.runInContext(code, ui);
  assert.equal(run('sharesInputRef({colinput: "B2", rowinput: "B2"})'), true);
  assert.equal(run('sharesInputRef({colinput: "B2", rowinput: "InputRate"})'), false);
  assert.equal(run('sharesInputRef({colinput: "", rowinput: ""})'), false);
  run('message = () => {}; calcDt(0); calcDt(1)');
  dialog.returnValue = 'no';
  run('resolveConfirmation()');
  assert.equal(run('SEM.confirmCallback'), null);
  assert.deepEqual(calls, [0]);

  run('calcDt(2)');
  dialog.returnValue = 'yes';
  run('resolveConfirmation()');
  assert.deepEqual(calls, [0, 2]);
  assert.equal(run('SEM.calculating'), 2);
  pending[1].onFailure(new Error('A calculation is running in another sidebar.'));
  assert.equal(run('SEM.calculating'), 1);
  pending[0].onSuccess();
  assert.equal(run('SEM.calculating'), 0);
});

test('manual focus resolves draft ranges and inputs without saving a DataTable', () => {
  const {ctx, saved, activatedRanges} = setup();
  ctx.focusDataTableRange({range: 'TableRange'}, 'range');
  ctx.focusDataTableRange({range: 'TableRange', colinput: 'B2'}, 'colinput');
  ctx.focusDataTableRange({range: 'TableRange', rowinput: 'InputRate'}, 'rowinput');
  ctx.focusDataTableRange({range: '', rowinput: "'Model'!C2"}, 'rowinput');
  assert.throws(() => ctx.focusDataTableRange({range: 'bad'}, 'range'), /valid range/);
  assert.deepEqual(activatedRanges, [[2, 4, 4], [2, 2, 2], [2, 2, 2], [1, 2, 3]]);
  assert.equal(saved(), '[]');
});
