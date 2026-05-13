require('./_setup');
const test = require('node:test');
const assert = require('node:assert');

const tools = require('../src/content/tools.js');

test('computeGeometry produces a bounding box and direction-preserving endpoints for a forward drag', () => {
  const g = tools.computeGeometry('line', { x: 10, y: 20 }, { x: 50, y: 80 });

  assert.strictEqual(g.x, 10);
  assert.strictEqual(g.y, 20);
  assert.strictEqual(g.width, 40);
  assert.strictEqual(g.height, 60);
  assert.strictEqual(g.x1, 0);
  assert.strictEqual(g.y1, 0);
  assert.strictEqual(g.x2, 40);
  assert.strictEqual(g.y2, 60);
});

test('computeGeometry preserves drag direction for a reverse drag (start > current)', () => {
  const g = tools.computeGeometry('line', { x: 50, y: 80 }, { x: 10, y: 20 });

  assert.strictEqual(g.x, 10);
  assert.strictEqual(g.y, 20);
  assert.strictEqual(g.width, 40);
  assert.strictEqual(g.height, 60);
  assert.strictEqual(g.x1, 40);
  assert.strictEqual(g.y1, 60);
  assert.strictEqual(g.x2, 0);
  assert.strictEqual(g.y2, 0);
});

test('computeGeometry for freehand uses the bounding box of all points and stores box-local path', () => {
  const points = [
    { x: 30, y: 50 },
    { x: 80, y: 100 },
    { x: 20, y: 90 },
  ];
  const g = tools.computeGeometry('freehand', { x: 30, y: 50 }, { x: 20, y: 90 }, points);

  assert.strictEqual(g.x, 20);
  assert.strictEqual(g.y, 50);
  assert.strictEqual(g.width, 60);
  assert.strictEqual(g.height, 50);
  assert.deepStrictEqual(g.freehandPath, [
    { x: 10, y: 0 },
    { x: 60, y: 50 },
    { x: 0, y: 40 },
  ]);
});

test('geometryFromAnnotation reconstructs box-local endpoints from stored startPoint/endPoint', () => {
  const annotation = {
    tool: 'arrow',
    x: 100,
    y: 200,
    width: 50,
    height: 30,
    startPoint: { x: 150, y: 230 },
    endPoint: { x: 100, y: 200 },
  };
  const g = tools.geometryFromAnnotation(annotation);

  assert.strictEqual(g.x, 100);
  assert.strictEqual(g.y, 200);
  assert.strictEqual(g.width, 50);
  assert.strictEqual(g.height, 30);
  assert.strictEqual(g.x1, 50);
  assert.strictEqual(g.y1, 30);
  assert.strictEqual(g.x2, 0);
  assert.strictEqual(g.y2, 0);
});

test('geometryFromAnnotation falls back to top-left → bottom-right when startPoint/endPoint missing', () => {
  const annotation = { tool: 'line', x: 0, y: 0, width: 80, height: 40 };
  const g = tools.geometryFromAnnotation(annotation);

  assert.strictEqual(g.x1, 0);
  assert.strictEqual(g.y1, 0);
  assert.strictEqual(g.x2, 80);
  assert.strictEqual(g.y2, 40);
});

test('geometryFromAnnotation passes through freehandPath for freehand tool', () => {
  const annotation = {
    tool: 'freehand',
    x: 10,
    y: 20,
    width: 30,
    height: 40,
    freehandPath: [{ x: 0, y: 0 }, { x: 30, y: 40 }],
  };
  const g = tools.geometryFromAnnotation(annotation);

  assert.deepStrictEqual(g.freehandPath, [{ x: 0, y: 0 }, { x: 30, y: 40 }]);
});

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

test('arrow.render produces an SVG path containing a shaft and two arrowhead lines', () => {
  const geometry = { width: 100, height: 0, x1: 0, y1: 0, x2: 100, y2: 0 };
  const svg = tools.arrow.render(geometry, '#0078d4');

  assert.strictEqual(svg.tagName.toLowerCase(), 'svg');
  const path = svg.querySelector('path');
  assert.ok(path, 'expected a <path> element');
  assert.strictEqual(path.getAttribute('stroke'), '#0078d4');
  assert.strictEqual(path.getAttribute('fill'), 'none');
  const d = path.getAttribute('d');
  assert.ok(d.includes('M 0 0'), 'path should start at x1,y1');
  assert.ok(d.includes('L 100 0'), 'path should draw shaft to x2,y2');
  assert.ok(d.match(/M [0-9.]+ [0-9.-]+ L 100 0 L [0-9.]+ [0-9.-]+/), 'path should include arrowhead');
});

test('circle.render produces an SVG ellipse sized to the bounding box', () => {
  const geometry = { width: 80, height: 40 };
  const svg = tools.circle.render(geometry, '#00ff00');

  const ellipse = svg.querySelector('ellipse');
  assert.ok(ellipse);
  assert.strictEqual(ellipse.getAttribute('cx'), '40');
  assert.strictEqual(ellipse.getAttribute('cy'), '20');
  assert.strictEqual(ellipse.getAttribute('rx'), '40');
  assert.strictEqual(ellipse.getAttribute('ry'), '20');
  assert.strictEqual(ellipse.getAttribute('stroke'), '#00ff00');
  assert.strictEqual(ellipse.getAttribute('fill'), 'none');
});

test('freehand.render produces an SVG path traversing the freehandPath points', () => {
  const geometry = {
    width: 50,
    height: 30,
    freehandPath: [{ x: 0, y: 0 }, { x: 25, y: 15 }, { x: 50, y: 30 }],
  };
  const svg = tools.freehand.render(geometry, '#ff00ff');

  const path = svg.querySelector('path');
  assert.ok(path);
  assert.strictEqual(path.getAttribute('stroke'), '#ff00ff');
  assert.strictEqual(path.getAttribute('fill'), 'none');
  assert.strictEqual(path.getAttribute('d'), 'M 0 0 L 25 15 L 50 30');
});

test('rectangle.render returns null because the box border is the rectangle', () => {
  const result = tools.rectangle.render({ width: 100, height: 50 }, '#000000');
  assert.strictEqual(result, null);
});
