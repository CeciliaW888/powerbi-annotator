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

const PowerBIAnnotatorTools = {
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
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = PowerBIAnnotatorTools;
}
