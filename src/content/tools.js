const SVG_NS = 'http://www.w3.org/2000/svg';

function createSvgContainer(width, height) {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('width', width);
  svg.setAttribute('height', height);
  svg.style.position = 'absolute';
  svg.style.top = '0';
  svg.style.left = '0';
  svg.style.pointerEvents = 'none';
  return svg;
}

function computeGeometry(toolName, startPoint, currentPoint, freehandPoints) {
  if (toolName === 'freehand' && freehandPoints && freehandPoints.length > 0) {
    const xs = freehandPoints.map(p => p.x);
    const ys = freehandPoints.map(p => p.y);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    const width = Math.max(...xs) - x;
    const height = Math.max(...ys) - y;
    return {
      x, y, width, height,
      freehandPath: freehandPoints.map(p => ({ x: p.x - x, y: p.y - y })),
    };
  }
  const x = Math.min(startPoint.x, currentPoint.x);
  const y = Math.min(startPoint.y, currentPoint.y);
  const width = Math.abs(currentPoint.x - startPoint.x);
  const height = Math.abs(currentPoint.y - startPoint.y);
  return {
    x, y, width, height,
    x1: startPoint.x - x,
    y1: startPoint.y - y,
    x2: currentPoint.x - x,
    y2: currentPoint.y - y,
  };
}

function geometryFromAnnotation(annotation) {
  const geometry = {
    x: annotation.x,
    y: annotation.y,
    width: annotation.width,
    height: annotation.height,
  };
  if (annotation.startPoint && annotation.endPoint) {
    geometry.x1 = annotation.startPoint.x - annotation.x;
    geometry.y1 = annotation.startPoint.y - annotation.y;
    geometry.x2 = annotation.endPoint.x - annotation.x;
    geometry.y2 = annotation.endPoint.y - annotation.y;
  } else {
    geometry.x1 = 0;
    geometry.y1 = 0;
    geometry.x2 = annotation.width;
    geometry.y2 = annotation.height;
  }
  if (annotation.freehandPath) {
    geometry.freehandPath = annotation.freehandPath;
  }
  return geometry;
}

const PowerBIAnnotatorTools = {
  computeGeometry,
  geometryFromAnnotation,
  rectangle: {
    name: 'rectangle',
    render() {
      return null;
    },
  },
  line: {
    name: 'line',
    render(geometry, color) {
      const svg = createSvgContainer(geometry.width, geometry.height);
      const line = document.createElementNS(SVG_NS, 'line');
      line.setAttribute('x1', geometry.x1);
      line.setAttribute('y1', geometry.y1);
      line.setAttribute('x2', geometry.x2);
      line.setAttribute('y2', geometry.y2);
      line.setAttribute('stroke', color);
      line.setAttribute('stroke-width', '3');
      svg.appendChild(line);
      return svg;
    },
  },
  arrow: {
    name: 'arrow',
    render(geometry, color) {
      const svg = createSvgContainer(geometry.width, geometry.height);
      const path = document.createElementNS(SVG_NS, 'path');
      const { x1, y1, x2, y2 } = geometry;
      const headLen = 20;
      const angle = Math.atan2(y2 - y1, x2 - x1);
      const ax1 = x2 - headLen * Math.cos(angle - Math.PI / 6);
      const ay1 = y2 - headLen * Math.sin(angle - Math.PI / 6);
      const ax2 = x2 - headLen * Math.cos(angle + Math.PI / 6);
      const ay2 = y2 - headLen * Math.sin(angle + Math.PI / 6);
      path.setAttribute('d', `M ${x1} ${y1} L ${x2} ${y2} M ${ax1} ${ay1} L ${x2} ${y2} L ${ax2} ${ay2}`);
      path.setAttribute('stroke', color);
      path.setAttribute('stroke-width', '3');
      path.setAttribute('fill', 'none');
      svg.appendChild(path);
      return svg;
    },
  },
  circle: {
    name: 'circle',
    render(geometry, color) {
      const svg = createSvgContainer(geometry.width, geometry.height);
      const ellipse = document.createElementNS(SVG_NS, 'ellipse');
      ellipse.setAttribute('cx', geometry.width / 2);
      ellipse.setAttribute('cy', geometry.height / 2);
      ellipse.setAttribute('rx', geometry.width / 2);
      ellipse.setAttribute('ry', geometry.height / 2);
      ellipse.setAttribute('stroke', color);
      ellipse.setAttribute('stroke-width', '3');
      ellipse.setAttribute('fill', 'none');
      svg.appendChild(ellipse);
      return svg;
    },
  },
  freehand: {
    name: 'freehand',
    render(geometry, color) {
      const svg = createSvgContainer(geometry.width, geometry.height);
      const path = document.createElementNS(SVG_NS, 'path');
      const d = geometry.freehandPath
        .map((p, i) => (i === 0 ? `M ${p.x} ${p.y}` : `L ${p.x} ${p.y}`))
        .join(' ');
      path.setAttribute('d', d);
      path.setAttribute('stroke', color);
      path.setAttribute('stroke-width', '3');
      path.setAttribute('fill', 'none');
      svg.appendChild(path);
      return svg;
    },
  },
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = PowerBIAnnotatorTools;
}
