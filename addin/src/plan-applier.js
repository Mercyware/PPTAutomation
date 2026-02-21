/* global PowerPoint */

const DEBUG_LOGS = (() => {
  if (typeof localStorage === "undefined") {
    return true;
  }
  const stored = localStorage.getItem("pptAutomationDebugLogs");
  if (stored === null) {
    return true;
  }
  return stored === "true";
})();

function debugLog(label, payload) {
  if (!DEBUG_LOGS) {
    return;
  }
  const timestamp = new Date().toISOString();
  // eslint-disable-next-line no-console
  console.log(`[plan-applier][${timestamp}] ${label}`, payload !== undefined ? payload : "");
}

async function applyExecutionPlan(plan, slideContext) {
  if (!plan || typeof plan !== "object" || !Array.isArray(plan.operations)) {
    throw new Error("Invalid plan payload");
  }

  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    if (!slides.items || slides.items.length === 0) {
      throw new Error("No slide found to apply the plan");
    }

    const activeSlide = await resolveActiveSlide(context, slides);
    if (!activeSlide) {
      throw new Error("Unable to resolve active slide");
    }

    const shapes = activeSlide.shapes;
    shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await context.sync();

    const shapeById = new Map();
    for (const shape of shapes.items) {
      shapeById.set(shape.id, shape);
    }

    const warnings = [];
    let appliedCount = 0;
    const usedShapeIds = new Set();

    // Track occupied regions so multiple inserts in a single plan don't collide with each other.
    // Seed from the provided slideContext; then append new insert bboxes as we create them.
    const occupiedRects = (Array.isArray(slideContext?.objects) ? slideContext.objects : [])
      .map((obj) => normalizeBbox(obj && obj.bbox))
      .filter((bbox) => bbox !== null);
    debugLog("Starting applyExecutionPlan", {
      operationCount: plan.operations.length,
      occupiedRectCount: occupiedRects.length,
    });
    for (const operation of plan.operations) {
      debugLog("Applying operation", operation);
      const result = await applyOperation({
        context,
        slide: activeSlide,
        operation,
        shapeById,
        selectedShapeIds: getSelectedShapeIdsFromContext(slideContext),
        slideContext,
        usedShapeIds,
        occupiedRects,
      });

      if (result.applied) {
        appliedCount += 1;
      }
      warnings.push(...result.warnings);
      debugLog("Operation result", result);
      if (result.occupiedBbox) {
        const rect = normalizeBbox([result.occupiedBbox.left, result.occupiedBbox.top, result.occupiedBbox.width, result.occupiedBbox.height]);
        if (rect) {
          occupiedRects.push(rect);
        }
      }
    }

    await context.sync();
    return {
      appliedCount,
      warnings: Array.from(new Set(warnings)),
    };
  });
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
    // fall through
  }

  return slides.items && slides.items.length > 0 ? slides.items[0] : null;
}

function getSelectedShapeIdsFromContext(slideContext) {
  const ids = slideContext && slideContext.selection && Array.isArray(slideContext.selection.shapeIds)
    ? slideContext.selection.shapeIds
    : [];
  return ids.filter((id) => typeof id === "string" && id.length > 0);
}

async function applyOperation({
  context,
  slide,
  operation,
  shapeById,
  selectedShapeIds,
  slideContext,
  usedShapeIds,
  occupiedRects,
}) {
  if (!operation || typeof operation !== "object") {
    return { applied: false, warnings: ["Skipped invalid operation"] };
  }

  const type = operation.type;
  if (!type) {
    return { applied: false, warnings: ["Skipped operation with missing type"] };
  }

  if (type === "delete") {
    debugLog("Skipping delete operation", operation);
    return { applied: false, warnings: ["Delete operations are not auto-applied in this phase"] };
  }

  const warnings = [];
  const tableRows = extractTableRows(operation);
  const imagePayload = extractImagePayload(operation);

  if (type === "insert" && imagePayload) {
    debugLog("Attempting image insert", imagePayload);
    const insertedImage = await tryInsertImage(
      context,
      slide,
      imagePayload,
      operation,
      slideContext,
      null,
      occupiedRects
    );
    if (insertedImage) {
      return { applied: true, warnings, occupiedBbox: insertedImage };
    }
    warnings.push("Image insert failed on this host; falling back to text rendering.");
  }

  if (type === "insert" && tableRows && tableRows.length > 0) {
    debugLog("Attempting table insert", { rowCount: tableRows.length, colCount: tableRows[0]?.length || 0 });
    const insertedTable = await tryInsertTable(
      context,
      slide,
      tableRows,
      operation,
      slideContext,
      null,
      occupiedRects
    );
    if (insertedTable) {
      return { applied: true, warnings, occupiedBbox: insertedTable };
    }
    warnings.push("Table insert is unavailable on this host; falling back to text rendering.");
  }

  const textContent = extractTextContent(operation);
  if (!textContent) {
    if (warnings.length > 0) {
      return { applied: false, warnings };
    }
    return { applied: false, warnings: ["Skipped operation without text content"] };
  }

  const subtitleIntent = type === "insert" && isSubtitleInsertIntent(operation, textContent);
  let preferredBbox = null;
  if (subtitleIntent) {
    const reserved = reserveSubtitlePlacement({
      shapeById,
      slideContext,
      text: textContent,
      occupiedRects,
    });
    if (reserved && reserved.bbox) {
      preferredBbox = reserved.bbox;
      if (Array.isArray(reserved.warnings) && reserved.warnings.length) {
        warnings.push(...reserved.warnings);
      }
    }
  }

  const intendedTarget = resolveIntendedExistingTarget({
    shapeById,
    operation,
    selectedShapeIds,
    slideContext,
  });

  if (type === "update" || type === "transform") {
    if (!intendedTarget) {
      return {
        applied: false,
        warnings: [
          `Skipped ${type}: intended target was not found. Provide a valid shape id or selection anchor.`,
        ],
      };
    }

      const written = trySetShapeText(intendedTarget, textContent, operation.styleBindings || {}, slideContext);
      if (written) {
        usedShapeIds.add(intendedTarget.id);
        return { applied: true, warnings };
      }

    return { applied: false, warnings: [...warnings, `Failed ${type}: target exists but is not writable`] };
  }

  const targetShape = intendedTarget;

  if (targetShape && type === "insert") {
    const safeToWrite = isSafeInsertTarget(targetShape, operation, slideContext);
    if (safeToWrite) {
      const written = trySetShapeText(targetShape, textContent, operation.styleBindings || {}, slideContext);
      if (written) {
        usedShapeIds.add(targetShape.id);
        return { applied: true, warnings };
      }
      preferredBbox = preferredBbox || getPreferredBboxForShape(targetShape, slideContext);
    }

    const alternativeTarget = resolveAlternativeInsertTarget({
      shapeById,
      selectedShapeIds,
      slideContext,
      usedShapeIds,
      excludedShapeIds: targetShape ? [targetShape.id] : [],
    });
    if (alternativeTarget) {
      const written = trySetShapeText(alternativeTarget, textContent, operation.styleBindings || {}, slideContext);
      if (written) {
        usedShapeIds.add(alternativeTarget.id);
        return { applied: true, warnings };
      }
      preferredBbox = preferredBbox || getPreferredBboxForShape(alternativeTarget, slideContext);
    }
  }

  const inserted = await tryInsertTextBox(
    context,
    slide,
    textContent,
    operation,
    slideContext,
    preferredBbox,
    occupiedRects
  );
  if (inserted) {
    return { applied: true, warnings, occupiedBbox: inserted };
  }

  return { applied: false, warnings: [...warnings, `Failed to apply ${type} operation`] };
}

function extractTextContent(operation) {
  if (!operation.content || typeof operation.content !== "object") {
    return "";
  }

  if (typeof operation.content.text === "string" && operation.content.text.trim()) {
    return normalizeEscapedNewLines(operation.content.text).trim();
  }

  if (Array.isArray(operation.content.rows) && operation.content.rows.length > 0) {
    const lines = operation.content.rows
      .filter((row) => Array.isArray(row))
      .map((row) => row.map((cell) => normalizeEscapedNewLines(String(cell))).join(" | "));
    return lines.join("\n").trim();
  }

  const tableRows = extractTableRows(operation);
  if (tableRows && tableRows.length > 0) {
    return tableRows.map((row) => row.join(" | ")).join("\n").trim();
  }

  const imagePayload = extractImagePayload(operation);
  if (imagePayload && imagePayload.url) {
    return `Image reference: ${imagePayload.url}`;
  }

  return "";
}

function extractTableRows(operation) {
  const content = operation && operation.content && typeof operation.content === "object" ? operation.content : null;
  if (!content) {
    return null;
  }

  const directRows = normalizeTableRows(content.rows);
  if (directRows && directRows.length > 0) {
    return directRows;
  }

  const table = content.table;
  if (!table) {
    return null;
  }

  if (Array.isArray(table)) {
    return normalizeTableRows(table);
  }

  if (typeof table === "object") {
    const rows = [];
    const headers = Array.isArray(table.headers) ? table.headers.map((cell) => toCellString(cell)) : [];
    if (headers.length > 0) {
      rows.push(headers);
    }

    const bodyRows = normalizeTableRows(table.rows || table.values);
    if (bodyRows && bodyRows.length > 0) {
      rows.push(...bodyRows);
    }

    return rows.length > 0 ? rows : null;
  }

  return null;
}

function normalizeTableRows(rows) {
  if (!Array.isArray(rows)) {
    return null;
  }

  const normalized = rows
    .filter((row) => Array.isArray(row))
    .map((row) => row.map((cell) => toCellString(cell)));

  if (normalized.length === 0) {
    return null;
  }

  const maxCols = normalized.reduce((max, row) => Math.max(max, row.length), 0);
  if (maxCols === 0) {
    return null;
  }

  return normalized.map((row) => {
    const padded = [...row];
    while (padded.length < maxCols) {
      padded.push("");
    }
    return padded;
  });
}

function toCellString(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return normalizeEscapedNewLines(String(value)).trim();
}

function extractImagePayload(operation) {
  const content = operation && operation.content && typeof operation.content === "object" ? operation.content : null;
  if (!content) {
    return null;
  }

  const image = content.image;
  if (typeof image === "string" && image.trim()) {
    return { url: image.trim(), alt: "" };
  }

  if (image && typeof image === "object") {
    const url = [image.url, image.src, image.dataUrl]
      .find((v) => typeof v === "string" && v.trim())
      || "";
    const base64 = typeof image.base64 === "string" && image.base64.trim() ? image.base64.trim() : "";
    const alt = typeof image.alt === "string" ? image.alt.trim() : "";
    if (url || base64) {
      return { url: url.trim(), base64, alt };
    }
  }

  if (typeof content.imageUrl === "string" && content.imageUrl.trim()) {
    return { url: content.imageUrl.trim(), alt: "" };
  }

  return null;
}

function normalizeEscapedNewLines(text) {
  let output = String(text || "");
  for (let i = 0; i < 3; i += 1) {
    output = output
      .replace(/\\\\r\\\\n/g, "\n")
      .replace(/\\\\n/g, "\n")
      .replace(/\\\\t/g, "\t")
      .replace(/\\r\\n/g, "\n")
      .replace(/\\n/g, "\n")
      .replace(/\\t/g, "\t");
  }
  return output;
}

function resolveTargetShape(shapeById, targetId) {
  if (typeof targetId !== "string" || !targetId.trim()) {
    return null;
  }
  return shapeById.get(targetId) || null;
}

function resolveIntendedExistingTarget({ shapeById, operation, selectedShapeIds, slideContext }) {
  const targeted = resolveTargetFromReference(shapeById, slideContext, operation?.target);
  if (targeted) {
    return targeted;
  }

  const anchor = operation?.anchor;
  if (anchor && String(anchor.strategy || "").toLowerCase() === "selection") {
    const selected = resolveTargetShapeFromSelection(shapeById, selectedShapeIds);
    if (selected) {
      return selected;
    }
  }

  const anchored = resolveTargetFromAnchor(shapeById, slideContext, anchor);
  if (anchored) {
    return anchored;
  }

  return null;
}

function resolveTargetFromReference(shapeById, slideContext, ref) {
  if (typeof ref !== "string" || !ref.trim()) {
    return null;
  }

  const direct = resolveTargetShape(shapeById, ref);
  if (direct) {
    return direct;
  }

  const numeric = Number(ref);
  if (Number.isInteger(numeric)) {
    return resolveTargetFromObjectIndex(shapeById, slideContext, numeric);
  }

  return null;
}

function resolveTargetFromAnchor(shapeById, slideContext, anchor) {
  if (!anchor || typeof anchor !== "object") {
    return null;
  }

  const ref = typeof anchor.ref === "string" ? anchor.ref.trim() : "";
  if (!ref) {
    return null;
  }

  return resolveTargetFromReference(shapeById, slideContext, ref);
}

function resolveTargetFromObjectIndex(shapeById, slideContext, numericRef) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const candidates = [numericRef, numericRef - 1];
  for (const index of candidates) {
    if (index >= 0 && index < objects.length) {
      const id = objects[index] && typeof objects[index].id === "string" ? objects[index].id : null;
      if (!id) {
        continue;
      }
      const shape = resolveTargetShape(shapeById, id);
      if (shape) {
        return shape;
      }
    }
  }
  return null;
}

function resolveBestPlaceholder(shapeById, slideContext, usedShapeIds, avoidUsed) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const objectById = new Map(objects.map((obj) => [obj.id, obj]));

  const candidates = [];
  for (const [id, shape] of shapeById.entries()) {
    if (avoidUsed && usedShapeIds.has(id)) {
      continue;
    }

    const obj = objectById.get(id);
    if (!obj) {
      continue;
    }

    const name = String(obj.name || shape.name || "").toLowerCase();
    const text = String(obj.text || "").trim();
    const isPlaceholderLike =
      name.includes("placeholder") ||
      name.includes("subtitle") ||
      name.includes("content") ||
      text.toLowerCase().includes("click to add");
    if (!isPlaceholderLike) {
      continue;
    }

    const bbox = normalizeBbox(obj.bbox) || normalizeBbox([shape.left, shape.top, shape.width, shape.height]);
    if (!bbox) {
      continue;
    }

    const area = bbox.width * bbox.height;
    const emptyBoost = isEmptyPlaceholderText(text) ? 1.8 : 1;
    const lowerHalfBoost = bbox.top > 220 ? 1.2 : 1;
    const score = area * emptyBoost * lowerHalfBoost;
    candidates.push({ shape, score });
  }

  candidates.sort((a, b) => b.score - a.score);
  return candidates.length > 0 ? candidates[0].shape : null;
}

function isSafeInsertTarget(shape, operation, slideContext) {
  const obj = getObjectForShape(slideContext, shape.id);
  const text = obj ? String(obj.text || "").trim() : "";
  const lowerName = String((obj && obj.name) || shape.name || "").toLowerCase();
  const anchorStrategy = String(operation?.anchor?.strategy || "").toLowerCase();

  // Never overwrite title placeholders for insert operations.
  if (lowerName.includes("title")) {
    return false;
  }

  // Safe when the shape is effectively empty placeholder text.
  if (!text || isEmptyPlaceholderText(text)) {
    return true;
  }

  // For selected/anchored inserts, do not replace existing authored text.
  if (anchorStrategy === "selection" || anchorStrategy === "placeholder") {
    return false;
  }

  return false;
}

function resolveAlternativeInsertTarget({
  shapeById,
  selectedShapeIds,
  slideContext,
  usedShapeIds,
  excludedShapeIds,
}) {
  const excluded = new Set(excludedShapeIds || []);
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const objectById = new Map(objects.map((obj) => [obj.id, obj]));

  const bySelection = resolveTargetShapeFromSelection(shapeById, selectedShapeIds);
  if (
    bySelection &&
    !excluded.has(bySelection.id) &&
    !usedShapeIds.has(bySelection.id) &&
    isSafeInsertTarget(bySelection, { anchor: { strategy: "selection" } }, slideContext)
  ) {
    return bySelection;
  }

  const placeholderCandidates = [];
  for (const [id, shape] of shapeById.entries()) {
    if (excluded.has(id) || usedShapeIds.has(id)) {
      continue;
    }
    const obj = objectById.get(id);
    const lowerName = String((obj && obj.name) || shape.name || "").toLowerCase();
    const placeholderLike =
      lowerName.includes("subtitle") ||
      lowerName.includes("content") ||
      lowerName.includes("placeholder");
    if (!placeholderLike) {
      continue;
    }
    if (!isSafeInsertTarget(shape, { anchor: { strategy: "placeholder" } }, slideContext)) {
      continue;
    }

    const bbox =
      normalizeBbox(obj && obj.bbox) || normalizeBbox([shape.left, shape.top, shape.width, shape.height]);
    if (!bbox) {
      continue;
    }

    const preferenceBoost =
      lowerName.includes("subtitle") || lowerName.includes("content") || lowerName.includes("placeholder")
        ? 1.4
        : 1;
    placeholderCandidates.push({
      shape,
      score: bbox.width * bbox.height * preferenceBoost,
    });
  }

  placeholderCandidates.sort((a, b) => b.score - a.score);
  return placeholderCandidates.length > 0 ? placeholderCandidates[0].shape : null;
}

function isEmptyPlaceholderText(text) {
  const t = String(text || "").toLowerCase();
  return t.includes("click to add") || t === "";
}

function getObjectForShape(slideContext, shapeId) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  return objects.find((obj) => obj.id === shapeId) || null;
}

function resolveTargetShapeFromSelection(shapeById, selectedShapeIds) {
  for (const id of selectedShapeIds) {
    const shape = shapeById.get(id);
    if (shape) {
      return shape;
    }
  }
  return null;
}

function resolveFirstTextCapableShape(shapeById, usedShapeIds, avoidUsed) {
  for (const shape of shapeById.values()) {
    if (avoidUsed && usedShapeIds && usedShapeIds.has(shape.id)) {
      continue;
    }
    if (shape && shape.textFrame && shape.textFrame.textRange) {
      return shape;
    }
  }
  return null;
}

function trySetShapeText(shape, text, styleBindings, slideContext) {
  try {
    if (!shape || !shape.textFrame || !shape.textFrame.textRange) {
      return false;
    }
    shape.textFrame.textRange.text = text;
    applyTextStyle(shape, text, styleBindings, slideContext);
    return true;
  } catch (_error) {
    return false;
  }
}

function applyTextStyle(shape, text, styleBindings, slideContext) {
  if (!shape || !shape.textFrame || !shape.textFrame.textRange) {
    return;
  }

  const font = shape.textFrame.textRange.font;
  const fontSize = estimateReadableFontSize(text, shape.width, shape.height);
  const subtitleLike = isSubtitleTextShape(shape, text, slideContext);
  if (font && Number.isFinite(fontSize)) {
    const adjustedSize = subtitleLike ? Math.min(26, Math.max(18, fontSize)) : fontSize;
    font.size = adjustedSize;
  }

  const resolved = resolveStyleBindings(styleBindings || {}, slideContext);
  if (font && resolved.font && typeof resolved.font === "string") {
    font.name = resolved.font;
  }
  if (font && resolved.color && typeof resolved.color === "string") {
    font.color = resolved.color;
  }

  // Best-effort text frame defaults (guarded because Office.js capabilities vary).
  // These improve readability without breaking older hosts.
  try {
    if (shape.textFrame && "wordWrap" in shape.textFrame) {
      shape.textFrame.wordWrap = true;
    }
  } catch (_error) {
    // ignore
  }

  try {
    // Some hosts support text frame margins; keep a consistent inner padding.
    const margin = subtitleLike ? 6 : 10;
    if (shape.textFrame && "marginLeft" in shape.textFrame) shape.textFrame.marginLeft = margin;
    if (shape.textFrame && "marginRight" in shape.textFrame) shape.textFrame.marginRight = margin;
    if (shape.textFrame && "marginTop" in shape.textFrame) shape.textFrame.marginTop = subtitleLike ? 4 : 8;
    if (shape.textFrame && "marginBottom" in shape.textFrame) shape.textFrame.marginBottom = subtitleLike ? 4 : 8;
  } catch (_error) {
    // ignore
  }
}

function isSubtitleTextShape(shape, text, slideContext) {
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const cleanText = String(text || "").trim();
  const lines = cleanText.split(/\r?\n/).filter((line) => line.trim().length > 0).length || 1;
  const nearTop = Number(shape?.top || 0) < slideHeight * 0.55;
  const compactBox = Number(shape?.height || 0) <= 110;
  const shortContent = cleanText.length <= 180 && lines <= 2;
  return nearTop && compactBox && shortContent;
}

function estimateReadableFontSize(text, width, height) {
  const cleanWidth = Math.max(120, Number(width || 0));
  const cleanHeight = Math.max(60, Number(height || 0));
  for (let size = 30; size >= 12; size -= 1) {
    if (fitsInBox(text, cleanWidth, cleanHeight, size)) {
      return size;
    }
  }
  return 12;
}

function fitsInBox(text, width, height, fontSize) {
  const innerWidth = Math.max(80, width - 20);
  const innerHeight = Math.max(40, height - 16);
  const charsPerLine = Math.max(10, Math.floor(innerWidth / (fontSize * 0.52)));
  const lines = countEstimatedLines(text, charsPerLine);
  const neededHeight = lines * fontSize * 1.28;
  return neededHeight <= innerHeight;
}

function countEstimatedLines(text, charsPerLine) {
  const parts = String(text || "").split(/\r?\n/);
  let lines = 0;
  for (const part of parts) {
    const len = Math.max(1, part.trim().length);
    lines += Math.ceil(len / charsPerLine);
  }
  return Math.max(1, lines);
}

function isSubtitleInsertIntent(operation, text) {
  const target = String(operation && operation.target ? operation.target : "").toLowerCase();
  const ref = String(operation && operation.anchor && operation.anchor.ref ? operation.anchor.ref : "").toLowerCase();
  const normalizedText = String(text || "").trim();
  const shortText = normalizedText.length > 0 && normalizedText.length <= 180;
  const mentionsSubtitle =
    target.includes("subtitle") ||
    ref.includes("subtitle") ||
    ref.includes("below-title") ||
    ref.includes("under-title");
  return mentionsSubtitle && shortText;
}

function reserveSubtitlePlacement({ shapeById, slideContext, text, occupiedRects }) {
  const draft = buildSubtitleDraft(slideContext, text, occupiedRects);
  if (!draft) {
    return null;
  }

  if (draft.availableGap >= draft.desiredHeight + 6) {
    return {
      bbox: {
        left: draft.left,
        top: draft.top,
        width: draft.width,
        height: draft.desiredHeight,
      },
      warnings: [],
    };
  }

  if (!Number.isFinite(draft.primaryTop)) {
    return {
      bbox: {
        left: draft.left,
        top: draft.top,
        width: draft.width,
        height: Math.max(44, Math.min(88, draft.desiredHeight)),
      },
      warnings: [],
    };
  }

  const neededDelta = draft.desiredHeight + 6 - Math.max(0, draft.availableGap);
  if (neededDelta <= 0) {
    return {
      bbox: {
        left: draft.left,
        top: draft.top,
        width: draft.width,
        height: draft.desiredHeight,
      },
      warnings: [],
    };
  }

  const shifted = shiftContentDownForSubtitle({
    shapeById,
    slideContext,
    startTop: draft.primaryTop,
    delta: neededDelta,
  });
  if (!shifted) {
    const compressedHeight = Math.max(44, Math.min(draft.desiredHeight, draft.availableGap - 6));
    if (compressedHeight >= 44) {
      return {
        bbox: {
          left: draft.left,
          top: draft.top,
          width: draft.width,
          height: compressedHeight,
        },
        warnings: ["Subtitle space was limited; applied compact subtitle placement."],
      };
    }
    return null;
  }

  rebuildOccupiedRectsFromShapes(shapeById, occupiedRects);
  return {
    bbox: {
      left: draft.left,
      top: draft.top,
      width: draft.width,
      height: draft.desiredHeight,
    },
    warnings: ["Shifted lower content to reserve subtitle space under the title."],
  };
}

function buildSubtitleDraft(slideContext, text, occupiedRects) {
  const slideWidth = Number(slideContext?.slide?.size?.w || 960);
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const margin = 24;
  const titleBox = detectTitleBox(slideContext, slideWidth, slideHeight);
  if (!titleBox) {
    return null;
  }

  const desiredWidth = Math.min(slideWidth - margin * 2, Math.max(420, Math.floor(slideWidth * 0.74)));
  const titleAlignedLeft = Number.isFinite(titleBox.left) ? titleBox.left : Math.floor((slideWidth - desiredWidth) / 2);
  const left = clamp(titleAlignedLeft, margin, Math.max(margin, slideWidth - desiredWidth - margin));
  const top = clamp(titleBox.bottom + 10, margin, Math.max(margin, slideHeight - 100 - margin));
  const desiredHeight = Math.max(48, Math.min(92, estimateHeightFromText(text, 22)));
  const primaryTop = findPrimaryContentTop(slideContext, occupiedRects, titleBox.bottom + 4);
  const availableGap = Number.isFinite(primaryTop) ? primaryTop - top : slideHeight - margin - top;

  return {
    left,
    top,
    width: desiredWidth,
    desiredHeight,
    availableGap,
    primaryTop,
  };
}

function findPrimaryContentTop(slideContext, occupiedRects, minTop) {
  const occupied =
    Array.isArray(occupiedRects) && occupiedRects.length
      ? occupiedRects
      : (Array.isArray(slideContext?.objects) ? slideContext.objects : [])
          .map((obj) => normalizeBbox(obj && obj.bbox))
          .filter((bbox) => bbox !== null);

  let bestTop = Number.POSITIVE_INFINITY;
  for (const rect of occupied) {
    if (!rect) continue;
    if (rect.top >= minTop && rect.top < bestTop && rect.height > 20 && rect.width > 120) {
      bestTop = rect.top;
    }
  }
  return bestTop;
}

function shiftContentDownForSubtitle({ shapeById, slideContext, startTop, delta }) {
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const margin = 20;
  const cleanDelta = Math.max(0, Math.ceil(Number(delta || 0)));
  if (!cleanDelta) {
    return true;
  }

  const movable = [];
  for (const shape of shapeById.values()) {
    const top = Number(shape.top);
    const height = Number(shape.height);
    if (!Number.isFinite(top) || !Number.isFinite(height)) {
      continue;
    }
    if (top >= startTop - 2) {
      if (top + cleanDelta + height > slideHeight - margin) {
        return false;
      }
      movable.push(shape);
    }
  }

  for (const shape of movable) {
    shape.top = Number(shape.top) + cleanDelta;
  }

  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  for (const obj of objects) {
    if (!obj || !Array.isArray(obj.bbox) || obj.bbox.length !== 4) {
      continue;
    }
    const top = Number(obj.bbox[1]);
    if (Number.isFinite(top) && top >= startTop - 2) {
      obj.bbox[1] = top + cleanDelta;
    }
  }

  return true;
}

function rebuildOccupiedRectsFromShapes(shapeById, occupiedRects) {
  if (!Array.isArray(occupiedRects)) {
    return;
  }
  occupiedRects.length = 0;
  for (const shape of shapeById.values()) {
    const rect = normalizeBbox([shape.left, shape.top, shape.width, shape.height]);
    if (rect) {
      occupiedRects.push(rect);
    }
  }
}

async function tryInsertTextBox(context, slide, text, operation, slideContext, preferredBbox, occupiedRects) {
  if (!slide || !slide.shapes || typeof slide.shapes.addTextBox !== "function") {
    return null;
  }

  const normalizedText = trimLongBulletText(text, slideContext);
  const bbox = getInsertBbox(operation, slideContext, normalizedText, preferredBbox, occupiedRects);
  try {
    const newShape = slide.shapes.addTextBox(normalizedText);
    if (newShape) {
      newShape.left = bbox.left;
      newShape.top = bbox.top;
      newShape.width = bbox.width;
      newShape.height = bbox.height;
      const rawBindings =
        operation && operation.styleBindings && typeof operation.styleBindings === "object" ? operation.styleBindings : {};
      const effectiveBindings = Object.keys(rawBindings).length ? rawBindings : { font: "theme.body" };
      applyTextStyle(newShape, normalizedText, effectiveBindings, slideContext);
    }
    await context.sync();
    return bbox;
  } catch (_error) {
    return null;
  }
}

async function tryInsertTable(context, slide, rows, operation, slideContext, preferredBbox, occupiedRects) {
  if (!slide || !slide.shapes || typeof slide.shapes.addTable !== "function") {
    debugLog("Table API unavailable on host");
    return null;
  }

  const textPreview = rows.map((row) => row.join(" | ")).join("\n");
  const bbox = getInsertBbox(operation, slideContext, textPreview, preferredBbox, occupiedRects);
  const rowCount = rows.length;
  const colCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  if (rowCount < 1 || colCount < 1) {
    return null;
  }

  try {
    const tableShape = slide.shapes.addTable(rowCount, colCount);
    debugLog("Table shape created", { rowCount, colCount });
    if (tableShape) {
      tableShape.left = bbox.left;
      tableShape.top = bbox.top;
      tableShape.width = bbox.width;
      tableShape.height = bbox.height;

      // Populate each cell after table creation. PowerPoint APIs can vary by host/version,
      // so we attempt both 0-based and 1-based cell indices.
      const table =
        typeof tableShape.getTable === "function"
          ? tableShape.getTable()
          : tableShape.table && typeof tableShape.table === "object"
            ? tableShape.table
            : null;
      if (table) {
        debugLog("Populating table cells");
        table.load("rowCount,columnCount");
        await context.sync();

        for (let r = 0; r < rows.length; r += 1) {
          for (let c = 0; c < rows[r].length; c += 1) {
            const value = rows[r][c];
            if (!value) {
              continue;
            }

            let cell = null;
            if (typeof table.getCellOrNullObject === "function") {
              cell = table.getCellOrNullObject(r, c);
              cell.load("isNullObject");
              await context.sync();
              if (cell.isNullObject) {
                cell = table.getCellOrNullObject(r + 1, c + 1);
                cell.load("isNullObject");
                await context.sync();
              }
            }

            if (cell && !cell.isNullObject) {
              cell.text = value;
            }
          }
        }
      } else {
        debugLog("Table object unavailable after shape creation");
      }
    }
    await context.sync();
    debugLog("Table insert success", bbox);
    return bbox;
  } catch (_error) {
    debugLog("Table insert failed", _error && _error.message ? _error.message : _error);
    return null;
  }
}

async function tryInsertImage(context, slide, imagePayload, operation, slideContext, preferredBbox, occupiedRects) {
  if (!slide || !slide.shapes) {
    return null;
  }

  const addImageFn =
    typeof slide.shapes.addImageFromBase64 === "function"
      ? "addImageFromBase64"
      : typeof slide.shapes.addImage === "function"
        ? "addImage"
        : null;

  if (!addImageFn) {
    debugLog("Image API unavailable on host");
    return null;
  }

  const base64 = await resolveImageBase64(imagePayload);
  if (!base64) {
    debugLog("Image payload could not be resolved to base64", imagePayload);
    return null;
  }

  const bbox = getInsertBbox(
    operation,
    slideContext,
    imagePayload.alt || "Image",
    preferredBbox,
    occupiedRects
  );

  try {
    const imageShape = slide.shapes[addImageFn](base64);
    if (imageShape) {
      imageShape.left = bbox.left;
      imageShape.top = bbox.top;
      imageShape.width = bbox.width;
      imageShape.height = bbox.height;
    }
    await context.sync();
    debugLog("Image insert success", bbox);
    return bbox;
  } catch (_error) {
    debugLog("Image insert failed", _error && _error.message ? _error.message : _error);
    return null;
  }
}

async function resolveImageBase64(imagePayload) {
  if (!imagePayload) {
    return null;
  }

  if (typeof imagePayload.base64 === "string" && imagePayload.base64.trim()) {
    return stripDataUrlPrefix(imagePayload.base64.trim());
  }

  if (!imagePayload.url) {
    return null;
  }

  const url = String(imagePayload.url).trim();
  if (url.startsWith("data:image/")) {
    return stripDataUrlPrefix(url);
  }

  try {
    const response = await fetch(url);
    if (!response.ok) {
      return null;
    }
    const blob = await response.blob();
    return await blobToBase64(blob);
  } catch (_error) {
    return null;
  }
}

function stripDataUrlPrefix(value) {
  const match = String(value || "").match(/^data:.*;base64,(.+)$/i);
  return match ? match[1] : String(value || "");
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Failed to read image payload"));
    reader.onloadend = () => {
      const result = typeof reader.result === "string" ? reader.result : "";
      resolve(stripDataUrlPrefix(result));
    };
    reader.readAsDataURL(blob);
  });
}

function getInsertBbox(operation, slideContext, text, preferredBbox, occupiedRects) {
  const hinted = normalizeBbox(preferredBbox);
  if (hinted && isReasonableBox(hinted, slideContext)) {
    return sanitizeBbox(stripRectBounds(hinted), slideContext);
  }

  if (
    operation &&
    operation.content &&
    operation.content.bbox &&
    Array.isArray(operation.content.bbox) &&
    operation.content.bbox.length === 4
  ) {
    const [left, top, width, height] = operation.content.bbox.map((v) => Number(v));
    const fromPlan = normalizeBbox([left, top, width, height]);
    if (fromPlan && isReasonableBox(fromPlan, slideContext) && !looksLikeCornerDefault(fromPlan)) {
      return sanitizeBbox({ left, top, width, height }, slideContext);
    }
  }

  const suggested = getBestFreeRegion(slideContext, text, hinted ? stripRectBounds(hinted) : null, occupiedRects);
  if (suggested) {
    return sanitizeBbox(suggested, slideContext);
  }

  const slideWidth = Number(slideContext?.slide?.size?.w || 960);
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const fallbackWidth = Math.max(520, slideWidth - 120);
  const fallbackHeight = Math.min(340, Math.max(190, estimateHeightFromText(text, 16)));
  return sanitizeBbox(
    {
      left: Math.max(30, Math.floor((slideWidth - fallbackWidth) / 2)),
      top: Math.max(120, Math.floor((slideHeight - fallbackHeight) / 2) + 20),
      width: fallbackWidth,
      height: fallbackHeight,
    },
    slideContext
  );
}

function getPreferredBboxForShape(shape, slideContext) {
  if (!shape) {
    return null;
  }

  const obj = getObjectForShape(slideContext, shape.id);
  const fromObj = normalizeBbox(obj && obj.bbox);
  if (fromObj) {
    return stripRectBounds(fromObj);
  }

  const fromShape = normalizeBbox([shape.left, shape.top, shape.width, shape.height]);
  if (fromShape) {
    return stripRectBounds(fromShape);
  }
  return null;
}

function getBestFreeRegion(slideContext, text, hintRect, occupiedRects) {
  const slideWidth = Number(slideContext?.slide?.size?.w || 960);
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const occupied =
    Array.isArray(occupiedRects) && occupiedRects.length
      ? occupiedRects
      : (Array.isArray(slideContext?.objects) ? slideContext.objects : [])
          .map((obj) => normalizeBbox(obj && obj.bbox))
          .filter((bbox) => bbox !== null);

  const margin = 24;
  const desiredHeight = Math.min(320, Math.max(120, estimateHeightFromText(text, 20)));
  const desiredWidth = Math.min(slideWidth - margin * 2, Math.max(420, Math.floor(slideWidth * 0.76)));
  const centerLeft = Math.max(margin, Math.floor((slideWidth - desiredWidth) / 2));

  const candidates = [];
  const titleBox = detectTitleBox(slideContext, slideWidth, slideHeight);

  if (titleBox) {
    candidates.push(
      withRectBounds({
        left: centerLeft,
        top: Math.min(slideHeight - desiredHeight - margin, titleBox.bottom + 18),
        width: desiredWidth,
        height: desiredHeight,
      })
    );
  }

  candidates.push(
    withRectBounds({
      left: centerLeft,
      top: Math.floor(slideHeight * 0.48),
      width: desiredWidth,
      height: Math.min(desiredHeight, Math.floor(slideHeight * 0.42)),
    })
  );

  const sortedByBottom = [...occupied].sort((a, b) => a.bottom - b.bottom);
  for (const box of sortedByBottom) {
    candidates.push(
      withRectBounds({
        left: centerLeft,
        top: box.bottom + 12,
        width: desiredWidth,
        height: Math.max(100, Math.min(slideHeight - (box.bottom + 12) - margin, desiredHeight)),
      })
    );
  }

  // Add a lightweight grid search around the center to find a low-overlap placement.
  // This helps when there are many objects and the heuristic candidates are all blocked.
  const xCandidates = uniqueNumbers([
    centerLeft,
    margin,
    Math.max(margin, Math.floor(slideWidth - desiredWidth - margin)),
    Math.max(margin, Math.floor(slideWidth * 0.12)),
    Math.max(margin, Math.floor(slideWidth * 0.18)),
  ]);
  const stepY = 12;
  const yStart = margin;
  const yEnd = Math.max(margin, Math.floor(slideHeight - desiredHeight - margin));
  for (let y = yStart; y <= yEnd; y += stepY) {
    for (const x of xCandidates) {
      candidates.push(
        withRectBounds({
          left: x,
          top: y,
          width: desiredWidth,
          height: desiredHeight,
        })
      );
    }
  }

  let best = null;
  let bestScore = Number.POSITIVE_INFINITY;
  for (const candidate of candidates) {
    if (!isInSlideBounds(candidate, slideWidth, slideHeight, margin)) {
      continue;
    }
    if (candidate.height < 90 || candidate.width < 260) {
      continue;
    }

    const overlapArea = occupied.reduce((sum, rect) => sum + overlapRectArea(candidate, rect), 0);
    const centerPenalty = Math.abs((candidate.left + candidate.width / 2) - slideWidth / 2);
    const preferredTop = resolvePreferredTop(slideContext, slideWidth, slideHeight);
    const verticalPenalty = Math.abs(candidate.top - preferredTop);
    const hintPenalty = hintRect ? rectDistance(candidate, withRectBounds(hintRect)) * 0.35 : 0;

    // Hard-penalize candidates that meaningfully overlap; prefer zero-overlap placements.
    const score = overlapArea * 2000 + centerPenalty * 1.8 + verticalPenalty * 1.1 + hintPenalty;

    if (score < bestScore) {
      bestScore = score;
      best = candidate;
    }
  }

  return best ? stripRectBounds(best) : null;
}

function estimateHeightFromText(text, fontSize) {
  const lines = countEstimatedLines(normalizeEscapedNewLines(text), 55);
  const lineHeight = Math.max(16, Number(fontSize || 16) * 1.35);
  return Math.ceil(lines * lineHeight + 24);
}

function trimLongBulletText(text, slideContext) {
  const normalized = normalizeEscapedNewLines(text);
  const lines = normalized.split(/\r?\n/).filter((line) => line.trim().length > 0);
  if (lines.length <= 22) {
    return normalized;
  }

  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const maxLines = slideHeight < 520 ? 16 : 20;
  const trimmed = lines.slice(0, maxLines);
  trimmed.push("...");
  return trimmed.join("\n");
}

function normalizeBbox(bbox) {
  if (!Array.isArray(bbox) || bbox.length !== 4) {
    return null;
  }
  const [left, top, width, height] = bbox.map((v) => Number(v));
  if (![left, top, width, height].every((v) => Number.isFinite(v))) {
    return null;
  }
  return {
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
  };
}

function rectsOverlap(a, b) {
  return !(a.right <= b.left || a.left >= b.right || a.bottom <= b.top || a.top >= b.bottom);
}

function overlapRectArea(a, b) {
  const x = Math.max(0, Math.min(a.right, b.right) - Math.max(a.left, b.left));
  const y = Math.max(0, Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top));
  return x * y;
}

function isInSlideBounds(rect, slideWidth, slideHeight, margin) {
  return (
    rect.left >= margin &&
    rect.top >= margin &&
    rect.right <= slideWidth - margin &&
    rect.bottom <= slideHeight - margin
  );
}

function detectTitleBox(slideContext, slideWidth, slideHeight) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const candidates = objects
    .map((obj) => {
      const bbox = normalizeBbox(obj.bbox);
      if (!bbox) {
        return null;
      }
      const name = String(obj.name || "").toLowerCase();
      const text = String(obj.text || "").trim();
      if (!text || isEmptyPlaceholderText(text)) {
        return null;
      }
      const likelyTitle =
        name.includes("title") ||
        (bbox.top < slideHeight * 0.38 && bbox.width > slideWidth * 0.38 && text.length < 120);
      if (!likelyTitle) {
        return null;
      }
      return bbox;
    })
    .filter((v) => v !== null);

  if (!candidates.length) {
    return null;
  }
  candidates.sort((a, b) => a.top - b.top);
  return candidates[0];
}

function isReasonableBox(box, slideContext) {
  const slideWidth = Number(slideContext?.slide?.size?.w || 960);
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  return box.width >= 240 && box.height >= 44 && box.right <= slideWidth + 2 && box.bottom <= slideHeight + 2;
}

function looksLikeCornerDefault(box) {
  return box.left <= 30 && box.top <= 40;
}

function withRectBounds(rect) {
  return {
    ...rect,
    right: rect.left + rect.width,
    bottom: rect.top + rect.height,
  };
}

function stripRectBounds(rect) {
  return {
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height,
  };
}

function resolveStyleBindings(styleBindings, slideContext) {
  const resolved = { font: null, color: null };
  const fonts = Array.isArray(slideContext?.themeHints?.fonts) ? slideContext.themeHints.fonts : [];
  const colors = Array.isArray(slideContext?.themeHints?.colors) ? slideContext.themeHints.colors : [];

  const rawFont = typeof styleBindings?.font === "string" ? styleBindings.font.trim() : "";
  const rawColor = typeof styleBindings?.color === "string" ? styleBindings.color.trim() : "";

  if (rawFont) {
    const lower = rawFont.toLowerCase();
    if (lower.startsWith("theme.")) {
      // Map theme tokens to an actual font seen on the slide.
      const titleFont = findLikelyTitleFont(slideContext) || (fonts.length ? fonts[0] : null);
      const bodyFont = findLikelyBodyFont(slideContext) || (fonts.length ? fonts[0] : null);
      resolved.font = lower.includes("title") ? titleFont : bodyFont;
    } else {
      resolved.font = rawFont;
    }
  }

  if (rawColor) {
    const lower = rawColor.toLowerCase();
    if (lower.startsWith("theme.")) {
      if (lower.includes("text")) {
        resolved.color = pickTextColor(colors, slideContext);
      } else {
        resolved.color = pickAccentColor(colors);
      }
    } else {
      resolved.color = normalizeColor(rawColor);
    }
  } else {
    // Avoid forcing a color if we don't have a strong signal; leave null to preserve theme.
    resolved.color = null;
  }

  return resolved;
}

function findLikelyTitleFont(slideContext) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  for (const obj of objects) {
    const name = String(obj?.name || "").toLowerCase();
    const font = obj?.style?.font;
    if (name.includes("title") && typeof font === "string" && font !== "unknown") {
      return font;
    }
  }
  return null;
}

function findLikelyBodyFont(slideContext) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const counts = new Map();
  for (const obj of objects) {
    const font = obj?.style?.font;
    if (typeof font !== "string" || !font.trim() || font === "unknown") continue;
    const name = String(obj?.name || "").toLowerCase();
    // De-emphasize titles when inferring body font.
    const boost = name.includes("title") ? 0.25 : 1;
    counts.set(font, (counts.get(font) || 0) + boost);
  }
  let best = null;
  let bestScore = 0;
  for (const [font, score] of counts.entries()) {
    if (score > bestScore) {
      best = font;
      bestScore = score;
    }
  }
  return best;
}

function pickAccentColor(colors) {
  const candidates = (Array.isArray(colors) ? colors : [])
    .map((c) => normalizeColor(c))
    .filter((c) => typeof c === "string");
  for (const c of candidates) {
    const lower = c.toLowerCase();
    if (lower !== "#000000" && lower !== "#ffffff") {
      return c;
    }
  }
  return candidates.length ? candidates[0] : null;
}

function pickTextColor(colors, slideContext) {
  const fromSlide = findLikelyBodyColor(slideContext);
  if (fromSlide) {
    return fromSlide;
  }

  const candidates = (Array.isArray(colors) ? colors : [])
    .map((c) => normalizeColor(c))
    .filter((c) => typeof c === "string");
  if (!candidates.length) {
    return null;
  }

  let best = candidates[0];
  let bestLum = colorLuminance(best);
  for (const c of candidates) {
    const lum = colorLuminance(c);
    if (lum < bestLum) {
      best = c;
      bestLum = lum;
    }
  }
  return best;
}

function findLikelyBodyColor(slideContext) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const counts = new Map();
  for (const obj of objects) {
    const raw = obj?.style?.color;
    const color = normalizeColor(raw);
    if (!color) continue;
    const name = String(obj?.name || "").toLowerCase();
    const boost = name.includes("title") ? 0.4 : 1;
    counts.set(color, (counts.get(color) || 0) + boost);
  }

  let best = null;
  let bestScore = -1;
  for (const [color, score] of counts.entries()) {
    const readabilityBonus = 1 - Math.min(1, colorLuminance(color));
    const total = score + readabilityBonus * 0.25;
    if (total > bestScore) {
      best = color;
      bestScore = total;
    }
  }
  return best;
}

function colorLuminance(color) {
  const normalized = normalizeColor(color);
  if (!normalized) return 1;
  const hex = normalized.slice(1);
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function normalizeColor(color) {
  const raw = String(color || "").trim();
  if (!raw) return null;
  const hex = raw.startsWith("#") ? raw.slice(1) : raw;
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  return `#${hex.toUpperCase()}`;
}

function sanitizeBbox(rect, slideContext) {
  const slideWidth = Number(slideContext?.slide?.size?.w || 960);
  const slideHeight = Number(slideContext?.slide?.size?.h || 540);
  const margin = 24;

  const minW = 260;
  const minH = 44;

  let left = Number(rect?.left || 0);
  let top = Number(rect?.top || 0);
  let width = Number(rect?.width || 0);
  let height = Number(rect?.height || 0);

  if (![left, top, width, height].every((v) => Number.isFinite(v))) {
    left = margin;
    top = Math.max(120, margin);
    width = Math.max(minW, slideWidth - margin * 2);
    height = Math.max(minH, Math.min(320, slideHeight - top - margin));
  }

  width = Math.max(minW, width);
  height = Math.max(minH, height);

  // Keep within slide bounds.
  left = clamp(left, margin, Math.max(margin, slideWidth - width - margin));
  top = clamp(top, margin, Math.max(margin, slideHeight - height - margin));

  // Snap to a small grid so placement looks intentional.
  const grid = 4;
  left = snapToGrid(left, grid);
  top = snapToGrid(top, grid);
  width = snapToGrid(width, grid);
  height = snapToGrid(height, grid);

  // Re-clamp after snapping.
  left = clamp(left, margin, Math.max(margin, slideWidth - width - margin));
  top = clamp(top, margin, Math.max(margin, slideHeight - height - margin));

  return { left, top, width, height };
}

function snapToGrid(value, grid) {
  const g = Math.max(1, Number(grid || 1));
  return Math.round(Number(value || 0) / g) * g;
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function uniqueNumbers(values) {
  const out = [];
  const seen = new Set();
  for (const v of values) {
    const n = Number(v);
    if (!Number.isFinite(n)) continue;
    const k = String(Math.round(n));
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(n);
  }
  return out;
}

function resolvePreferredTop(slideContext, slideWidth, slideHeight) {
  const titleBox = detectTitleBox(slideContext, slideWidth, slideHeight);
  if (titleBox) {
    return Math.min(slideHeight - 160, Math.max(80, titleBox.bottom + 18));
  }
  return Math.floor(slideHeight * 0.45);
}

function rectDistance(a, b) {
  // Distance between rectangles (0 if overlapping).
  const ax1 = a.left, ay1 = a.top, ax2 = a.right, ay2 = a.bottom;
  const bx1 = b.left, by1 = b.top, bx2 = b.right, by2 = b.bottom;

  const dx = ax2 < bx1 ? bx1 - ax2 : bx2 < ax1 ? ax1 - bx2 : 0;
  const dy = ay2 < by1 ? by1 - ay2 : by2 < ay1 ? ay1 - by2 : 0;
  return Math.sqrt(dx * dx + dy * dy);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    applyExecutionPlan,
  };
}

if (typeof window !== "undefined") {
  window.PPTAutomation = window.PPTAutomation || {};
  window.PPTAutomation.applyExecutionPlan = applyExecutionPlan;
}
