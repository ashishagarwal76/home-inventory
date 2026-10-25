const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

const appRoot = path.join(__dirname, '..', 'apps-script');
const schemaContext = vm.createContext({});
new vm.Script(fs.readFileSync(path.join(appRoot, 'Schema.gs'), 'utf8')).runInContext(schemaContext);
const schema = JSON.parse(JSON.stringify(schemaContext.inventoryTemplateSchema_()));
assert.equal(schema.length, 9, 'The reusable core schema has nine tables.');
const sources = ['Configuration.gs', 'Setup.gs'].map(file =>
  fs.readFileSync(path.join(appRoot, file), 'utf8')).join('\n');

const clone = value => JSON.parse(JSON.stringify(value));
const present = value => value !== undefined && value !== null && value !== '';
function display(value) {
  if (value && typeof value === 'object' && 'formula' in value) return String(value.value || '');
  return value == null ? '' : String(value);
}

function harness(options = {}) {
  const state = {
    properties: { ...(options.properties || {}) },
    sheets: new Map(), events: [], logs: [], formCalls: [],
    lockHeld: false, waits: 0, releases: 0, photoCalls: 0
  };
  function sheet(name, input = {}) {
    const object = {
      name,
      data: clone(input.data || []),
      maxColumns: input.maxColumns || 26,
      getMaxColumns() { return this.maxColumns; },
      getDataRange() {
        return { isBlank: () => !this.data.some(row => row.some(present)) };
      },
      getRange(row, column, rowCount, columnCount) {
        assert.ok(columnCount > 0 && column + columnCount - 1 <= this.maxColumns);
        const target = this;
        return {
          getDisplayValues() {
            return Array.from({ length: rowCount }, (_, offset) =>
              Array.from({ length: columnCount }, (_, index) =>
                display(target.data[row - 1 + offset]?.[column - 1 + index])));
          },
          setValues(values) {
            if (options.writeError) throw options.writeError;
            assert.equal(values.length, rowCount);
            assert.ok(values.every(valuesRow => valuesRow.length === columnCount));
            state.events.push(['headers', name]);
            values.forEach((valuesRow, offset) => {
              const targetRow = target.data[row - 1 + offset] ||= [];
              valuesRow.forEach((value, index) => { targetRow[column - 1 + index] = value; });
            });
          }
        };
      },
      insertColumnsAfter(position, count) {
        assert.equal(position, this.maxColumns);
        state.events.push(['columns', name]);
        this.maxColumns += count;
      }
    };
    return object;
  }
  for (const [name, input] of Object.entries(options.sheets || {})) state.sheets.set(name, sheet(name, input));
  const spreadsheet = {
    getId: () => 'new-owner-spreadsheet',
    getSheetByName: name => state.sheets.get(name) || null,
    insertSheet(name) {
      if (options.insertError) throw options.insertError;
      assert.ok(!state.sheets.has(name));
      state.events.push(['insert', name]);
      const created = sheet(name);
      state.sheets.set(name, created);
      return created;
    }
  };
  const lock = {
    waitLock(timeout) {
      state.waits++;
      assert.equal(timeout, 30000);
      if (options.waitError) throw options.waitError;
      assert.equal(state.lockHeld, false, 'The script lock must not be acquired twice.');
      state.lockHeld = true;
    },
    releaseLock() {
      state.releases++;
      assert.equal(state.lockHeld, true);
      state.lockHeld = false;
      if (options.releaseError) throw options.releaseError;
    }
  };
  const properties = {
    getProperty: key => Object.hasOwn(state.properties, key) ? state.properties[key] : null,
    setProperty(key, value) {
      if (options.propertyError) throw options.propertyError;
      state.events.push(['property', key]);
      state.properties[key] = value;
    }
  };
  const context = vm.createContext({
    console: { log: message => state.logs.push(String(message)) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => options.active === false ? null : spreadsheet,
      flush() {
        state.events.push(['flush']);
        if (options.flushError) throw options.flushError;
      }
    },
    LockService: { getScriptLock: () => lock },
    PropertiesService: { getScriptProperties: () => properties },
    FormApp: {
      openById(id) {
        state.formCalls.push(id);
        if (options.formError) throw options.formError;
        return { formID: id };
      }
    },
    inventoryTemplateSchema_: () => clone(options.schema || schema),
    setupPhotoStorage() {
      assert.equal(state.lockHeld, false, 'Photo setup must run after the setup lock is released.');
      state.photoCalls++;
      state.events.push(['photo']);
      if (options.photoError) throw options.photoError;
    }
  });
  new vm.Script(sources).runInContext(context);
  return { context, state };
}

const compatibleSheets = () => Object.fromEntries(schema.map(table => [table.name, { data: [table.headers] }]));
const mutations = state => state.events.filter(event => ['insert', 'headers', 'columns', 'property'].includes(event[0]));

test('fresh bound setup writes the supplied nine header rows, then stores configuration and runs photos', () => {
  const { context, state } = harness();
  assert.match(context.setupInventory(), /setup is complete/i);
  assert.equal(state.sheets.size, 9);
  for (const table of schema) assert.deepEqual(state.sheets.get(table.name).data, [table.headers]);
  assert.equal(state.properties.INVENTORY_SPREADSHEET_ID, 'new-owner-spreadsheet');
  assert.equal(state.photoCalls, 1);
  assert.equal(state.waits, state.releases);
  assert.equal(state.formCalls.length, 0);
  assert.ok(state.events.findIndex(event => event[0] === 'property') > state.events.findLastIndex(event => event[0] === 'headers'));
  assert.equal(state.events.at(-1)[0], 'photo');
});

test('rerunning preserves records, formulas, extra columns, other properties and unrelated tabs', () => {
  const sheets = compatibleSheets();
  sheets[schema[0].name].data[0].push('Owner custom column');
  sheets[schema[0].name].data.push(['existing-record', { formula: '=1+1', value: 2 }, 'saved details']);
  sheets.PersonalNotes = { data: [['Do not change'], ['Private note']] };
  const { context, state } = harness({ sheets, properties: {
    INVENTORY_SPREADSHEET_ID: 'new-owner-spreadsheet', LAST_ITEM_NUMBER: '42'
  } });
  const before = clone([...state.sheets].map(([name, value]) => [name, value.data]));
  context.setupInventory();
  context.setupInventory();
  assert.deepEqual([...state.sheets].map(([name, value]) => [name, value.data]), before);
  assert.deepEqual(mutations(state), []);
  assert.equal(state.properties.LAST_ITEM_NUMBER, '42');
  assert.equal(state.photoCalls, 2);
});

test('a blank existing tab is reused and expanded without creating sample records', () => {
  const table = schema.find(entry => entry.headers.length > 1);
  const { context, state } = harness({ sheets: { [table.name]: { data: [], maxColumns: 1 } } });
  const original = state.sheets.get(table.name);
  context.setupInventory();
  assert.equal(state.sheets.get(table.name), original);
  assert.deepEqual(original.data, [table.headers]);
  assert.equal(original.maxColumns, table.headers.length);
  assert.equal(state.events.filter(event => event[0] === 'insert').length, 8);
});

test('late header mismatch prevents every creation, header write and property change', () => {
  const last = schema.at(-1);
  const { context, state } = harness({ sheets: {
    [schema[0].name]: { data: [] },
    [last.name]: { data: [['Wrong header'], ['preserve this record']] }
  } });
  const before = clone([...state.sheets].map(([name, value]) => [name, value.data]));
  assert.throws(() => context.setupInventory(), /column 1 must be/);
  assert.deepEqual(mutations(state), []);
  assert.deepEqual([...state.sheets].map(([name, value]) => [name, value.data]), before);
  assert.equal(state.photoCalls, 0);
  assert.equal(state.releases, 1);
});

test('blank-looking formulas and data below an empty header are not treated as an empty tab', () => {
  for (const data of [[[{ formula: '=""', value: '' }]], [[], ['existing record']]]) {
    const { context, state } = harness({ sheets: { [schema[0].name]: { data } } });
    assert.throws(() => context.setupInventory(), /must be/);
    assert.deepEqual(mutations(state), []);
  }
});

test('a populated tab narrower than its required headers fails without expanding it', () => {
  const table = schema.find(entry => entry.headers.length > 1);
  const { context, state } = harness({ sheets: {
    [table.name]: { data: [[table.headers[0]], ['saved']], maxColumns: 1 }
  } });
  assert.throws(() => context.setupInventory(), /column 2 must be/);
  assert.deepEqual(mutations(state), []);
});

test('a conflicting spreadsheet configuration fails before changing tabs or properties', () => {
  const { context, state } = harness({ properties: { INVENTORY_SPREADSHEET_ID: 'different-sheet' } });
  assert.throws(() => context.setupInventory(), /different spreadsheet/);
  assert.deepEqual(mutations(state), []);
  assert.equal(state.properties.INVENTORY_SPREADSHEET_ID, 'different-sheet');
  assert.equal(state.photoCalls, 0);
  assert.equal(state.releases, 1);
});

test('setup requires the bound spreadsheet context before acquiring the lock', () => {
  const { context, state } = harness({ active: false });
  assert.throws(() => context.setupInventory(), /Extensions > Apps Script/);
  assert.equal(state.waits, 0);
  assert.deepEqual(mutations(state), []);
});

test('empty or malformed spreadsheet configuration produces a clear setup error', () => {
  for (const value of [undefined, '', '  ', 'https://example.test/sheet', 'wrong id']) {
    const properties = value === undefined ? {} : { INVENTORY_SPREADSHEET_ID: value };
    const { context, state } = harness({ properties });
    assert.throws(() => context.inventoryConfiguredSpreadsheetID_(), /setupInventory/);
    assert.deepEqual(mutations(state), []);
  }
  const { context } = harness({ properties: { INVENTORY_SPREADSHEET_ID: '  own-sheet_1  ' } });
  assert.equal(context.inventoryConfiguredSpreadsheetID_(), 'own-sheet_1');
});

test('scanner configuration is optional and trimmed without embedding a default host', () => {
  assert.equal(harness().context.inventoryScannerUrl_(), '');
  const { context } = harness({ properties: { INVENTORY_SCANNER_URL: '  https://scanner.example.test/  ' } });
  assert.equal(context.inventoryScannerUrl_(), 'https://scanner.example.test/');
});

test('optional form configuration opens only a known configured form', () => {
  const { context, state } = harness({ properties: { INVENTORY_ADD_FORM_ID: '  own-form_1  ' } });
  assert.equal(context.inventoryOptionalForm_('INVENTORY_CONSUME_FORM_ID'), null);
  assert.deepEqual(context.inventoryOptionalForm_('INVENTORY_ADD_FORM_ID'), { formID: 'own-form_1' });
  assert.deepEqual(state.formCalls, ['own-form_1']);
  assert.throws(() => context.inventoryOptionalForm_('UNRELATED_PROPERTY'), /Unknown inventory Form setting/);
  assert.deepEqual(state.formCalls, ['own-form_1']);
  const invalid = harness({ properties: { INVENTORY_ADD_FORM_ID: 'https://example.test/form' } });
  assert.throws(() => invalid.context.inventoryOptionalForm_('INVENTORY_ADD_FORM_ID'), /Google Form ID/);
  assert.deepEqual(invalid.state.formCalls, []);
});

test('Google Form access errors propagate without mutating configuration', () => {
  const error = new Error('No access to form');
  const { context, state } = harness({ properties: { INVENTORY_ADD_FORM_ID: 'own-form' }, formError: error });
  assert.throws(() => context.inventoryOptionalForm_('INVENTORY_ADD_FORM_ID'), candidate => candidate === error);
  assert.deepEqual(mutations(state), []);
});

test('invalid template schema is rejected before any spreadsheet change', () => {
  for (const badSchema of [[], [{ name: 'Items', headers: ['Duplicate', 'Duplicate'] }]]) {
    const { context, state } = harness({ schema: badSchema });
    assert.throws(() => context.setupInventory(), /schema is (missing|invalid)/);
    assert.deepEqual(mutations(state), []);
    assert.equal(state.releases, 1);
  }
});

test('lock acquisition failure does not release an unacquired lock or start photo setup', () => {
  const error = new Error('Lock timeout');
  const { context, state } = harness({ waitError: error });
  assert.throws(() => context.setupInventory(), candidate => candidate === error);
  assert.equal(state.releases, 0);
  assert.equal(state.photoCalls, 0);
  assert.deepEqual(mutations(state), []);
});

test('header/configuration/flush failures always release the acquired lock and skip photos', () => {
  for (const key of ['insertError', 'writeError', 'propertyError', 'flushError']) {
    const error = new Error(key);
    const { context, state } = harness({ [key]: error });
    assert.throws(() => context.setupInventory(), candidate => candidate === error);
    assert.equal(state.lockHeld, false);
    assert.equal(state.releases, 1);
    assert.equal(state.photoCalls, 0);
    assert.equal(state.logs.length, 0);
  }
});

test('release errors and photo errors are surfaced without reporting successful setup', () => {
  const releaseError = new Error('Release failed');
  const first = harness({ releaseError });
  assert.throws(() => first.context.setupInventory(), candidate => candidate === releaseError);
  assert.equal(first.state.photoCalls, 0);
  assert.equal(first.state.logs.length, 0);
  const photoError = new Error('Photo folder not accessible');
  const second = harness({ photoError });
  assert.throws(() => second.context.setupInventory(), candidate => candidate === photoError);
  assert.equal(second.state.lockHeld, false);
  assert.equal(second.state.releases, 1);
  assert.equal(second.state.photoCalls, 1);
  assert.equal(second.state.logs.length, 0);
  assert.equal(second.state.properties.INVENTORY_SPREADSHEET_ID, 'new-owner-spreadsheet');
});
