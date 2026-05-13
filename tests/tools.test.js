require('./_setup');
const test = require('node:test');
const assert = require('node:assert');

const tools = require('../src/content/tools.js');

test('line.render returns an SVG containing a <line> element with the given color and endpoints', () => {
  const geometry = { width: 100, height: 50, x1: 0, y1: 0, x2: 100, y2: 50 };
  const svg = tools.line.render(geometry, '#ff0000');

  assert.strictEqual(svg.tagName.toLowerCase(), 'svg');
  assert.strictEqual(svg.getAttribute('width'), '100');
  assert.strictEqual(svg.getAttribute('height'), '50');

  const lineEl = svg.querySelector('line');
  assert.ok(lineEl, 'expected an inner <line> element');
  assert.strictEqual(lineEl.getAttribute('stroke'), '#ff0000');
  assert.strictEqual(lineEl.getAttribute('x1'), '0');
  assert.strictEqual(lineEl.getAttribute('y1'), '0');
  assert.strictEqual(lineEl.getAttribute('x2'), '100');
  assert.strictEqual(lineEl.getAttribute('y2'), '50');
});
