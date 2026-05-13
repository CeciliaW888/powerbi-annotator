require('./_setup');
const test = require('node:test');
const assert = require('node:assert');

const { createPageStore } = require('../src/content/page-store.js');

function fakeStorage(initial = {}) {
  let store = JSON.parse(JSON.stringify(initial));
  return {
    get(keys, cb) { cb(JSON.parse(JSON.stringify(store))); },
    set(obj, cb) { Object.assign(store, obj); if (cb) cb(); },
    _peek: () => JSON.parse(JSON.stringify(store)),
  };
}

function locationOf(pathname, search = '') {
  return () => ({ pathname, search });
}

test('current() derives key from reportId + ReportSection hash when both present in path', () => {
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: locationOf('/groups/abc/reports/report-123/ReportSectionXYZ'),
  });
  const page = store.current();
  assert.strictEqual(page.key, 'report-123#ReportSectionXYZ');
  assert.strictEqual(page.reportId, 'report-123');
  assert.strictEqual(page.sectionHash, 'ReportSectionXYZ');
});

test('current() falls back to pathname+search when no reportId/sectionHash can be derived', () => {
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: locationOf('/some/random/path', '?foo=bar'),
  });
  const page = store.current();
  assert.strictEqual(page.key, '/some/random/path?foo=bar');
  assert.strictEqual(page.sectionHash, null);
});

test('current() extracts reportId from ?reportId= query param when path has none', () => {
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: locationOf('/reportEmbed', '?reportId=qid-456&pageName=ReportSectionABC'),
  });
  const page = store.current();
  assert.strictEqual(page.reportId, 'qid-456');
});

test('current() returns the display name resolved at call time', () => {
  let name = 'Sales Overview';
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
    displayNameResolver: () => name,
  });
  assert.strictEqual(store.current().displayName, 'Sales Overview');
  name = 'Margins';
  assert.strictEqual(store.current().displayName, 'Margins');
});

test('init() loads annotations from storage and exposes them on current()', async () => {
  const store = createPageStore({
    storage: fakeStorage({
      annotations: { 'r1#ReportSectionAAA': [{ id: 1, comment: 'hello' }] },
    }),
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
  });
  await store.init();
  const page = store.current();
  assert.strictEqual(page.annotations.length, 1);
  assert.strictEqual(page.annotations[0].comment, 'hello');
});

test('saveAnnotations() persists for the current page and writes through to storage', async () => {
  const storage = fakeStorage();
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
  });
  await store.init();
  store.saveAnnotations([{ id: 1, comment: 'first' }]);
  assert.deepStrictEqual(storage._peek().annotations, {
    'r1#ReportSectionAAA': [{ id: 1, comment: 'first' }],
  });
});

test('saveAnnotations() does not clobber other pages annotations', async () => {
  const storage = fakeStorage({
    annotations: { 'r1#OldPage': [{ id: 9, comment: 'old' }] },
  });
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
  });
  await store.init();
  store.saveAnnotations([{ id: 1, comment: 'new' }]);
  const data = storage._peek().annotations;
  assert.strictEqual(data['r1#OldPage'][0].comment, 'old');
  assert.strictEqual(data['r1#ReportSectionAAA'][0].comment, 'new');
});

test('list() returns annotated pages in the order given by pageOrderResolver', async () => {
  const storage = fakeStorage({
    annotations: {
      'r1#PageB': [{ id: 1 }],
      'r1#PageA': [{ id: 2 }, { id: 3 }],
      'r1#PageC': [{ id: 4 }],
    },
  });
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
    pageOrderResolver: () => ['r1#PageA', 'r1#PageB', 'r1#PageC'],
  });
  await store.init();
  const pages = store.list();
  assert.deepStrictEqual(pages.map(p => p.key), ['r1#PageA', 'r1#PageB', 'r1#PageC']);
  assert.strictEqual(pages[0].annotations.length, 2);
});

test('list() puts pages not in the report order at the end', async () => {
  const storage = fakeStorage({
    annotations: {
      'r1#KnownPage': [{ id: 1 }],
      'r1#UnknownPage': [{ id: 2 }],
    },
  });
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
    pageOrderResolver: () => ['r1#KnownPage'],
  });
  await store.init();
  const pages = store.list();
  assert.deepStrictEqual(pages.map(p => p.key), ['r1#KnownPage', 'r1#UnknownPage']);
});

test('onPageChange fires when checkPageChange detects the URL changed', async () => {
  let path = '/groups/g/reports/r1/ReportSectionAAA';
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: () => ({ pathname: path, search: '' }),
  });
  await store.init();
  const seen = [];
  store.onPageChange((page) => seen.push(page.key));

  path = '/groups/g/reports/r1/ReportSectionBBB';
  store.checkPageChange();

  assert.strictEqual(seen.length, 1);
  assert.strictEqual(seen[0], 'r1#ReportSectionBBB');
});

test('onPageChange does not fire when URL is unchanged', async () => {
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
  });
  await store.init();
  const seen = [];
  store.onPageChange(() => seen.push('fired'));
  store.checkPageChange();
  store.checkPageChange();
  assert.strictEqual(seen.length, 0);
});

test('migrates legacy pathname-only key to canonical key on init', async () => {
  const storage = fakeStorage({
    annotations: { '/groups/g/reports/r1/ReportSectionAAA': [{ id: 99 }] },
  });
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionAAA'),
  });
  await store.init();
  const data = storage._peek().annotations;
  assert.ok(data['r1#ReportSectionAAA'], 'expected canonical key after migration');
  assert.strictEqual(data['r1#ReportSectionAAA'][0].id, 99);
  assert.strictEqual(data['/groups/g/reports/r1/ReportSectionAAA'], undefined);
});

test('deleteAnnotations(key) removes that pages entry from storage', async () => {
  const storage = fakeStorage({
    annotations: {
      'r1#A': [{ id: 1 }],
      'r1#B': [{ id: 2 }],
    },
  });
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionA'),
  });
  await store.init();
  store.deleteAnnotations('r1#B');
  assert.deepStrictEqual(Object.keys(storage._peek().annotations), ['r1#A']);
});

test('deleteAll() wipes everything from storage', async () => {
  const storage = fakeStorage({
    annotations: { 'r1#A': [{ id: 1 }], 'r1#B': [{ id: 2 }] },
  });
  const store = createPageStore({
    storage,
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionA'),
  });
  await store.init();
  store.deleteAll();
  assert.deepStrictEqual(storage._peek().annotations, {});
});

test('onDataChange fires after saveAnnotations', async () => {
  const store = createPageStore({
    storage: fakeStorage(),
    locationProvider: locationOf('/groups/g/reports/r1/ReportSectionA'),
  });
  await store.init();
  let fired = 0;
  store.onDataChange(() => fired++);
  store.saveAnnotations([{ id: 1 }]);
  assert.strictEqual(fired, 1);
});
