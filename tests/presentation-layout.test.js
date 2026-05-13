const test = require('node:test');
const assert = require('node:assert');

const layout = require('../src/content/presentation-layout.js');

test('SLIDE_DIMENSIONS exposes widescreen layout constants', () => {
  const d = layout.SLIDE_DIMENSIONS;
  assert.strictEqual(d.width, 13.33);
  assert.strictEqual(d.height, 7.5);
  assert.ok(d.margin > 0);
  assert.ok(d.titleHeight > 0);
});

test('computeImageFit returns the source dims unchanged when source already fits inside box', () => {
  const fit = layout.computeImageFit(100, 50, 200, 200);
  assert.strictEqual(fit.width, 200);
  assert.strictEqual(fit.height, 100);
});

test('computeImageFit scales by width when image is wider relative to box', () => {
  // image aspect 2:1, box aspect 1:1 → fit by width
  const fit = layout.computeImageFit(200, 100, 100, 100);
  assert.strictEqual(fit.width, 100);
  assert.strictEqual(fit.height, 50);
});

test('computeImageFit scales by height when image is taller relative to box', () => {
  // image aspect 1:2, box aspect 1:1 → fit by height
  const fit = layout.computeImageFit(100, 200, 100, 100);
  assert.strictEqual(fit.width, 50);
  assert.strictEqual(fit.height, 100);
});

test('computeImageFit centers the result inside the box (offsetX/offsetY)', () => {
  const fit = layout.computeImageFit(200, 100, 100, 100); // becomes 100x50
  assert.strictEqual(fit.offsetX, 0);
  assert.strictEqual(fit.offsetY, 25); // centered vertically: (100-50)/2
});

test('chunkComments splits a list across slides when it exceeds the per-slide max', () => {
  const comments = Array.from({ length: 25 }, (_, i) => ({ id: i + 1, text: `c${i}` }));
  const chunks = layout.chunkComments(comments, 10);
  assert.strictEqual(chunks.length, 3);
  assert.strictEqual(chunks[0].length, 10);
  assert.strictEqual(chunks[1].length, 10);
  assert.strictEqual(chunks[2].length, 5);
});

test('chunkComments returns a single chunk when comments fit on one slide', () => {
  const comments = [{ id: 1 }, { id: 2 }];
  const chunks = layout.chunkComments(comments, 10);
  assert.deepStrictEqual(chunks, [comments]);
});

test('chunkComments returns no chunks for an empty input', () => {
  assert.deepStrictEqual(layout.chunkComments([], 10), []);
});

test('commentsPerSlide returns a positive integer count based on available height and line height', () => {
  const n = layout.commentsPerSlide({ availableHeight: 6, lineHeight: 0.28 });
  assert.ok(n >= 1);
  assert.strictEqual(n, Math.floor(6 / 0.28));
});
