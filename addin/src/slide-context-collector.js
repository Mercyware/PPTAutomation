/* global PowerPoint, Office */

/**
 * Collects a normalized slide context payload from the active/selected PowerPoint slide.
 * Uses best-effort extraction so the payload remains usable across host capability levels.
 */
async function collectSlideContext() {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const slides = presentation.slides;
    slides.load("items");
    await context.sync();

    const emptyContext = {
      slide: {},
      selection: { shapeIds: [] },
      themeHints: {},
      objects: [],
    };

    if (!slides.items || slides.items.length === 0) {
      return emptyContext;
    }

    const selectedShapeIds = await getSelectedShapeIds(context);
    const activeSlide = await resolveActiveSlide(context, slides);
    if (!activeSlide) {
      return emptyContext;
    }

    activeSlide.load("id");
    const shapes = activeSlide.shapes;
    shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await context.sync();

    const objects = [];
    const themeHints = {
      fonts: new Set(),
      colors: new Set(),
    };

    for (const shape of shapes.items) {
      const shapeType = toShapeType(shape.type);
      const object = {
        id: shape.id,
        name: typeof shape.name === "string" ? shape.name : "",
        type: shapeType,
        bbox: [shape.left, shape.top, shape.width, shape.height],
      };

      const textMeta = await extractTextForShape(context, shape);
      if (textMeta) {
        object.text = textMeta.text;
        object.style = textMeta.style;
        if (textMeta.style.font && textMeta.style.font !== "unknown") {
          themeHints.fonts.add(textMeta.style.font);
        }
        if (textMeta.style.color && textMeta.style.color !== "unknown") {
          themeHints.colors.add(textMeta.style.color);
        }
      }

      if (isTableShape(shapeType)) {
        const tableMeta = await extractTableForShape(context, shape);
        if (tableMeta) {
          object.table = tableMeta;
        }
      }

      if (isChartShape(shapeType)) {
        object.chart = {
          detected: true,
          type: "unknown",
        };
      }

      objects.push(object);
    }

    const inferredSize = inferSlideSize(objects);
    const rawSlide = await collectRawSlidePayloads();

    return {
      slide: {
        id: activeSlide.id,
        size: inferredSize,
      },
      selection: {
        shapeIds: selectedShapeIds,
      },
      themeHints: {
        fonts: Array.from(themeHints.fonts).slice(0, 8),
        colors: Array.from(themeHints.colors).slice(0, 12),
      },
      objects,
      rawSlide,
    };
  });
}

async function collectRawSlidePayloads() {
  const result = {
    ooxml: null,
    imageBase64: null,
    exportedAt: new Date().toISOString(),
  };

  try {
    const imageData = await getSelectedDataAsyncSafe(Office?.CoercionType?.Image);
    if (typeof imageData === "string" && imageData.trim()) {
      result.imageBase64 = stripDataUrlPrefix(imageData);
    } else if (imageData && typeof imageData === "object") {
      const candidate = imageData.value || imageData.image || imageData.data;
      if (typeof candidate === "string" && candidate.trim()) {
        result.imageBase64 = stripDataUrlPrefix(candidate);
      }
    }
  } catch (_error) {
    // best-effort only
  }

  return result;
}

function getSelectedDataAsyncSafe(coercionType) {
  if (!coercionType) {
    return Promise.resolve(null);
  }

  if (!Office || !Office.context || !Office.context.document) {
    return Promise.resolve(null);
  }

  if (typeof Office.context.document.getSelectedDataAsync !== "function") {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(coercionType, (asyncResult) => {
      if (!asyncResult || asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        resolve(null);
        return;
      }
      resolve(asyncResult.value || null);
    });
  });
}

function stripDataUrlPrefix(value) {
  const text = String(value || "");
  const match = text.match(/^data:.*;base64,(.+)$/i);
  return match ? match[1] : text;
}

async function resolveActiveSlide(context, slides) {
  try {
    if (typeof context.presentation.getSelectedSlides === "function") {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items/id");
      await context.sync();
      if (selectedSlides.items && selectedSlides.items.length > 0) {
        return selectedSlides.items[0];
      }
    }
  } catch (_error) {
    // Fall through to first available slide.
  }

  return slides.items && slides.items.length > 0 ? slides.items[0] : null;
}

async function getSelectedShapeIds(context) {
  try {
    if (typeof context.presentation.getSelectedShapes !== "function") {
      return [];
    }

    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items/id");
    await context.sync();

    if (!selectedShapes.items) {
      return [];
    }

    return selectedShapes.items
      .map((shape) => shape.id)
      .filter((id) => typeof id === "string" && id.length > 0);
  } catch (_error) {
    return [];
  }
}

async function extractTextForShape(context, shape) {
  try {
    if (!shape.textFrame || !shape.textFrame.textRange) {
      return null;
    }

    shape.textFrame.textRange.load("text,font/name,font/color,font/size");
    await context.sync();

    const textRange = shape.textFrame.textRange;
    const text = typeof textRange.text === "string" ? textRange.text.trim() : "";
    if (!text) {
      return null;
    }

    return {
      text: text.slice(0, 2000),
      style: {
        font: textRange.font && typeof textRange.font.name === "string" ? textRange.font.name : "unknown",
        color:
          textRange.font && typeof textRange.font.color === "string"
            ? textRange.font.color
            : "unknown",
        size:
          textRange.font && typeof textRange.font.size === "number"
            ? textRange.font.size
            : null,
      },
    };
  } catch (_error) {
    return null;
  }
}

async function extractTableForShape(context, shape) {
  try {
    if (!shape.table) {
      return null;
    }

    shape.table.load("rowCount,columnCount");
    await context.sync();

    const rowCount = Number(shape.table.rowCount || 0);
    const columnCount = Number(shape.table.columnCount || 0);

    return {
      rowCount,
      columnCount,
      values: [],
    };
  } catch (_error) {
    return null;
  }
}

function toShapeType(value) {
  if (typeof value === "string") {
    return value;
  }
  if (value === null || value === undefined) {
    return "unknown";
  }
  return String(value);
}

function isTableShape(type) {
  return typeof type === "string" && type.toLowerCase().includes("table");
}

function isChartShape(type) {
  return typeof type === "string" && type.toLowerCase().includes("chart");
}

function inferSlideSize(objects) {
  const bboxes = objects
    .map((obj) => obj.bbox)
    .filter((bbox) => Array.isArray(bbox) && bbox.length === 4)
    .map((bbox) => bbox.map((v) => Number(v)))
    .filter((vals) => vals.every((v) => Number.isFinite(v)));

  if (!bboxes.length) {
    return { w: 960, h: 540 };
  }

  let maxRight = 960;
  let maxBottom = 540;
  for (const [left, top, width, height] of bboxes) {
    maxRight = Math.max(maxRight, left + width);
    maxBottom = Math.max(maxBottom, top + height);
  }

  return { w: Math.round(maxRight), h: Math.round(maxBottom) };
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    collectSlideContext,
  };
}

if (typeof window !== "undefined") {
  window.PPTAutomation = window.PPTAutomation || {};
  window.PPTAutomation.collectSlideContext = collectSlideContext;
}
