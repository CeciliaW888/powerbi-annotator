const SLIDE_DIMENSIONS = (() => {
  const width = 13.33;
  const height = 7.5;
  const margin = 0.3;
  const titleHeight = 0.4;
  const titleY = margin;
  const contentY = titleY + titleHeight + 0.2;
  const contentHeight = height - contentY - margin;
  const contentWidth = width - margin * 2;
  const screenshotWidthRatio = 0.80;
  const commentsWidthRatio = 0.18;
  return {
    width,
    height,
    margin,
    titleHeight,
    titleY,
    contentY,
    contentHeight,
    contentWidth,
    screenshotWidth: contentWidth * screenshotWidthRatio,
    commentsWidth: contentWidth * commentsWidthRatio,
    commentsX:
      margin + contentWidth * screenshotWidthRatio +
      (contentWidth - contentWidth * screenshotWidthRatio - contentWidth * commentsWidthRatio),
    commentLineHeight: 0.28,
  };
})();

function computeImageFit(srcWidth, srcHeight, boxWidth, boxHeight) {
  const srcAspect = srcWidth / srcHeight;
  const boxAspect = boxWidth / boxHeight;
  let width, height;
  if (srcAspect > boxAspect) {
    width = boxWidth;
    height = boxWidth / srcAspect;
  } else {
    height = boxHeight;
    width = boxHeight * srcAspect;
  }
  return {
    width,
    height,
    offsetX: (boxWidth - width) / 2,
    offsetY: (boxHeight - height) / 2,
  };
}

function chunkComments(comments, maxPerSlide) {
  if (!comments || comments.length === 0) return [];
  const chunks = [];
  for (let i = 0; i < comments.length; i += maxPerSlide) {
    chunks.push(comments.slice(i, i + maxPerSlide));
  }
  return chunks;
}

function commentsPerSlide({ availableHeight, lineHeight }) {
  return Math.floor(availableHeight / lineHeight);
}

const api = { SLIDE_DIMENSIONS, computeImageFit, chunkComments, commentsPerSlide };

if (typeof window !== 'undefined') {
  window.PowerBIAnnotatorPresentationLayout = api;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = api;
}
