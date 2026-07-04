(function () {
  function getCanvasPageRect(canvasEl, win) {
    const r = canvasEl.getBoundingClientRect();
    const w = win || window;
    return {
      left: r.left + (w.scrollX || 0),
      top: r.top + (w.scrollY || 0),
      width: r.width,
      height: r.height,
    };
  }

  function ptToRel(pt, c) {
    return pt ? { x: (pt.x - c.left) / c.width, y: (pt.y - c.top) / c.height } : null;
  }

  function ptToAbs(pt, c) {
    return pt ? { x: c.left + pt.x * c.width, y: c.top + pt.y * c.height } : null;
  }

  function annotationToRelative(annotation, canvasRect) {
    const c = canvasRect;
    return Object.assign({}, annotation, {
      coordSpace: 'canvas',
      rel: {
        x: (annotation.x - c.left) / c.width,
        y: (annotation.y - c.top) / c.height,
        w: annotation.width / c.width,
        h: annotation.height / c.height,
      },
      relStart: ptToRel(annotation.startPoint, c),
      relEnd: ptToRel(annotation.endPoint, c),
      relFreehand: annotation.freehandPath
        ? annotation.freehandPath.map((p) => ptToRel(p, c))
        : null,
    });
  }

  function annotationToAbsolute(annotation, canvasRect) {
    const c = canvasRect;
    return Object.assign({}, annotation, {
      x: c.left + annotation.rel.x * c.width,
      y: c.top + annotation.rel.y * c.height,
      width: annotation.rel.w * c.width,
      height: annotation.rel.h * c.height,
      startPoint: ptToAbs(annotation.relStart, c),
      endPoint: ptToAbs(annotation.relEnd, c),
      freehandPath: annotation.relFreehand
        ? annotation.relFreehand.map((p) => ptToAbs(p, c))
        : null,
    });
  }

  function migrateAnnotation(annotation, canvasRect) {
    if (annotation.coordSpace === 'canvas' && annotation.rel) return annotation;
    return annotationToRelative(annotation, canvasRect);
  }

  const api = { getCanvasPageRect, annotationToRelative, annotationToAbsolute, migrateAnnotation };
  if (typeof window !== 'undefined') window.PowerBIAnnotatorCoords = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})();
