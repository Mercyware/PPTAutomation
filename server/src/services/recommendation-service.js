const { generateStructuredJson } = require("./ollama-client");
const { buildRecommendationPrompt, buildPlanPrompt } = require("./prompts");
const { tryParseJson } = require("../utils/json");
const { validateRecommendations, ALLOWED_OUTPUT_TYPES } = require("../validation/recommendations");
const { validateExecutionPlan } = require("../validation/execution-plan");
const { resolvePresentationImage } = require("./image-search-service");

const DEBUG_LOGS = process.env.DEBUG_LOGS === "true" || process.env.NODE_ENV !== "production";
const DESIGN_HINT_KEYWORDS = [
  "theme",
  "style",
  "contrast",
  "hierarchy",
  "spacing",
  "alignment",
  "readability",
  "whitespace",
  "visual",
  "polish",
];

function debugLog(label, payload) {
  if (!DEBUG_LOGS) {
    return;
  }
  const timestamp = new Date().toISOString();
  console.log(`[recommendation-service][${timestamp}] ${label}`);
  if (payload !== undefined) {
    try {
      console.log(
        typeof payload === "string" ? payload : JSON.stringify(payload, null, 2)
      );
    } catch (_error) {
      console.log(String(payload));
    }
  }
}

function summarizeContext(slideContext) {
  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const objectSummary = objects.map((obj) => ({
    id: obj.id,
    name: obj.name,
    type: obj.type,
    text: typeof obj.text === "string" ? obj.text : undefined,
    style: obj.style,
    table: obj.table,
    chart: obj.chart,
    bbox: obj.bbox,
  }));

  const rawSlide =
    slideContext?.rawSlide && typeof slideContext.rawSlide === "object"
      ? {
          ooxml:
            typeof slideContext.rawSlide.ooxml === "string" ? slideContext.rawSlide.ooxml : null,
          imageBase64:
            typeof slideContext.rawSlide.imageBase64 === "string"
              ? slideContext.rawSlide.imageBase64
              : null,
          exportedAt: slideContext.rawSlide.exportedAt || null,
        }
      : null;

  return {
    slide: slideContext?.slide || {},
    selection: slideContext?.selection || {},
    themeHints: slideContext?.themeHints || {},
    objectCount: objects.length,
    objects: objectSummary,
    rawSlide,
  };
}

function fallbackRecommendations(userPrompt) {
  return [
    {
      id: "improve-layout",
      title: "Improve slide layout hierarchy",
      description: "Refine spacing, alignment, and visual hierarchy for better readability.",
      outputType: "layout-improvement",
      confidence: 0.81,
      applyHints: ["align-to-grid", "increase-whitespace", "preserve-theme"],
    },
    {
      id: "summarize-content",
      title: "Summarize this slide",
      description: "Create a concise summary with key points.",
      outputType: "summary",
      confidence: 0.78,
      applyHints: ["reuse-title-placeholder", "preserve-theme"],
    },
    {
      id: "convert-to-smartart",
      title: "Convert to SmartArt",
      description: "Restructure core points into an editable SmartArt-style process visual.",
      outputType: "smartart",
      confidence: 0.76,
      applyHints: ["prefer-smartart", "preserve-theme", "avoid-overlap"],
    },
    {
      id: "next-steps",
      title: "Generate next steps",
      description: `Create actionable next steps for: "${userPrompt.slice(0, 80)}"`,
      outputType: "list",
      confidence: 0.74,
      applyHints: ["bullet-list", "keep-editable"],
    },
  ];
}

function cleanSentence(value, maxLen) {
  const text = String(value || "")
    .replace(/\s+/g, " ")
    .replace(/^[-*]\s+/, "")
    .trim();
  return text.slice(0, maxLen);
}

function toKebabId(value, fallback) {
  const raw = cleanSentence(value, 120).toLowerCase();
  const normalized = raw
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return normalized || fallback;
}

function canonicalizeOutputType(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) {
    return "other";
  }

  const compact = raw.replace(/[\s_-]+/g, "");
  if (compact === "layoutimprovement" || compact === "layout") {
    return "layout-improvement";
  }
  if (compact === "bulletlist" || compact === "bullets") {
    return "list";
  }
  if (compact === "smartart" || compact === "smartdiagram" || compact === "diagram") {
    return "smartart";
  }
  return raw;
}

function toHintToken(value) {
  return cleanSentence(value, 80)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizeApplyHints(applyHints, outputType) {
  const source = Array.isArray(applyHints) ? applyHints : [];
  const defaults = defaultHintsForType(outputType);
  const combined = source.length > 0 ? source : defaults;

  const normalized = [];
  const seen = new Set();
  for (const hint of combined) {
    const token = toHintToken(hint);
    if (!token || seen.has(token)) {
      continue;
    }
    seen.add(token);
    normalized.push(token);
  }

  if (!seen.has("preserve-theme")) {
    normalized.push("preserve-theme");
    seen.add("preserve-theme");
  }
  if (!seen.has("avoid-overlap")) {
    normalized.push("avoid-overlap");
    seen.add("avoid-overlap");
  }

  if ((outputType === "layout-improvement" || outputType === "smartart") && !seen.has("improve-visual-hierarchy")) {
    normalized.push("improve-visual-hierarchy");
  }
  if (outputType === "image" && !seen.has("prefer-raster-image-url")) {
    normalized.push("prefer-raster-image-url");
  }

  return normalized.slice(0, 8);
}

function includesAnyKeyword(text, keywords) {
  const source = String(text || "").toLowerCase();
  return keywords.some((keyword) => source.includes(keyword));
}

function isDesignFocusedRecommendation(item) {
  if (item.outputType === "layout-improvement" || item.outputType === "smartart") {
    return true;
  }

  const hints = Array.isArray(item.applyHints) ? item.applyHints : [];
  if (hints.some((hint) => includesAnyKeyword(hint, DESIGN_HINT_KEYWORDS))) {
    return true;
  }

  return includesAnyKeyword(item.description, DESIGN_HINT_KEYWORDS);
}

function recommendationPresentationPriority(item) {
  const outputTypeWeights = {
    "layout-improvement": 0.22,
    smartart: 0.2,
    chart: 0.15,
    image: 0.12,
    table: 0.1,
    summary: 0.08,
    list: 0.06,
    other: 0.04,
  };

  let score = Number(item.confidence || 0);
  score += outputTypeWeights[item.outputType] || outputTypeWeights.other;

  const hints = Array.isArray(item.applyHints) ? item.applyHints : [];
  for (const hint of hints) {
    if (includesAnyKeyword(hint, DESIGN_HINT_KEYWORDS)) {
      score += 0.03;
    }
    if (String(hint).includes("preserve-theme")) {
      score += 0.04;
    }
  }

  if (includesAnyKeyword(item.description, DESIGN_HINT_KEYWORDS)) {
    score += 0.05;
  }

  return score;
}

function isSmartArtType(outputType) {
  return outputType === "smartart" || outputType === "diagram";
}

function defaultHintsForType(outputType) {
  switch (outputType) {
    case "table":
      return ["prefer-table-shape", "preserve-theme", "avoid-overlap"];
    case "image":
      return ["insert-image-shape", "prefer-raster-image-url", "preserve-theme", "avoid-overlap"];
    case "chart":
      return ["prefer-chart-shape", "preserve-theme", "avoid-overlap"];
    case "smartart":
    case "diagram":
      return ["prefer-smartart", "keep-editable", "preserve-theme", "avoid-overlap"];
    case "summary":
      return ["reuse-title-placeholder", "keep-editable", "preserve-theme"];
    default:
      return ["preserve-theme", "avoid-overlap", "keep-editable"];
  }
}

function defaultDescriptionForType(outputType) {
  switch (outputType) {
    case "table":
      return "Reorganize existing content into a clear editable table with meaningful headers.";
    case "chart":
      return "Transform numeric or categorical content into a chart to make comparisons easier.";
    case "image":
      return "Add a relevant visual to support the slide message without disrupting layout.";
    case "smartart":
    case "diagram":
      return "Turn the content into an editable SmartArt-style visual to clarify flow and structure.";
    case "summary":
      return "Condense slide content into concise key points for faster understanding.";
    case "layout-improvement":
      return "Improve alignment, spacing, and hierarchy to increase clarity and readability.";
    default:
      return "Apply a focused improvement that preserves design and keeps content editable.";
  }
}

function isFormattingRecommendation(item) {
  return item.outputType === "layout-improvement";
}

function isCompletionRecommendation(item) {
  return item.outputType !== "layout-improvement";
}

function ensureCoverageCandidates(items) {
  const enhanced = [...items];

  const hasFormatting = enhanced.some((item) => isFormattingRecommendation(item));
  const hasCompletion = enhanced.some((item) => isCompletionRecommendation(item));
  const hasDesignFocus = enhanced.some((item) => isDesignFocusedRecommendation(item));

  if (!hasFormatting) {
    enhanced.push({
      id: "format-for-readability",
      title: "Improve formatting readability",
      description:
        "Refine spacing, font hierarchy, and alignment so the slide is easier to scan and present.",
      outputType: "layout-improvement",
      confidence: 0.86,
      applyHints: ["align-to-grid", "increase-whitespace", "preserve-theme"],
    });
  }

  if (!hasCompletion) {
    enhanced.push({
      id: "complete-core-content",
      title: "Complete the core content",
      description:
        "Fill missing placeholders with concise, audience-ready content aligned to the inferred slide intent.",
      outputType: "list",
      confidence: 0.87,
      applyHints: ["fill-placeholders", "preserve-theme", "keep-editable"],
    });
  }

  if (!hasDesignFocus) {
    enhanced.push({
      id: "theme-visual-polish",
      title: "Theme and visual polish pass",
      description:
        "Harmonize fonts, colors, spacing, and contrast so the slide looks presentation-ready and on-brand.",
      outputType: "layout-improvement",
      confidence: 0.89,
      applyHints: ["preserve-theme", "improve-visual-hierarchy", "increase-whitespace", "improve-contrast"],
    });
  }

  return enhanced;
}

function enforceRecommendationQuality(items) {
  const deduped = [];
  const seenKeys = new Set();

  for (const item of items) {
    const key = `${item.outputType}::${item.title.toLowerCase()}`;
    if (seenKeys.has(key)) {
      continue;
    }
    seenKeys.add(key);
    deduped.push(item);
  }

  const byPriority = deduped.sort((a, b) => {
    const scoreDiff = recommendationPresentationPriority(b) - recommendationPresentationPriority(a);
    if (Math.abs(scoreDiff) > 0.0001) {
      return scoreDiff;
    }
    return b.confidence - a.confidence;
  });

  const selected = [];
  const usedIds = new Set();

  const formatting = byPriority.find((item) => isFormattingRecommendation(item));
  if (formatting) {
    selected.push(formatting);
    usedIds.add(formatting.id);
  }

  const completion = byPriority.find(
    (item) => isCompletionRecommendation(item) && !usedIds.has(item.id)
  );
  if (completion) {
    selected.push(completion);
    usedIds.add(completion.id);
  }

  const usedTypes = new Set(selected.map((item) => item.outputType));

  for (const item of byPriority) {
    if (usedIds.has(item.id)) {
      continue;
    }
    if (!usedTypes.has(item.outputType)) {
      selected.push(item);
      usedIds.add(item.id);
      usedTypes.add(item.outputType);
    }
    if (selected.length >= 4) {
      break;
    }
  }

  if (selected.length < 3) {
    for (const item of byPriority) {
      if (!usedIds.has(item.id)) {
        selected.push(item);
        usedIds.add(item.id);
      }
      if (selected.length >= 3) {
        break;
      }
    }
  }

  if (selected.length < 3) {
    return [];
  }

  return selected.slice(0, 6).map((item, index) => ({
    ...item,
    id: item.id || `rec-${index + 1}`,
  }));
}

function normalizeRecommendations(raw) {
  const recommendations = Array.isArray(raw?.recommendations) ? raw.recommendations : [];
  const normalized = recommendations
    .map((item, index) => ({
      id:
        typeof item.id === "string" && item.id.trim()
          ? toKebabId(item.id.trim(), `rec-${index + 1}`)
          : `rec-${index + 1}`,
      title: typeof item.title === "string" ? item.title.trim() : "Recommendation",
      description: typeof item.description === "string" ? item.description.trim() : "",
      outputType: canonicalizeOutputType(item.outputType),
      confidence:
        typeof item.confidence === "number"
          ? Math.max(0, Math.min(1, item.confidence))
          : 0.5,
      applyHints: Array.isArray(item.applyHints) ? item.applyHints.filter((v) => typeof v === "string") : [],
    }))
    .map((item) => ({
      ...item,
      outputType: ALLOWED_OUTPUT_TYPES.has(item.outputType) ? item.outputType : "other",
      description:
        cleanSentence(item.description, 300) || defaultDescriptionForType(item.outputType),
      title: cleanSentence(item.title, 120),
      applyHints: normalizeApplyHints(item.applyHints, item.outputType),
    }))
    .filter((item) => item.title.length > 0)
    .map((item) => ({
      ...item,
      id: toKebabId(item.id || item.title, "rec"),
      confidence: Math.max(0.35, Math.min(0.98, item.confidence)),
    }));

  const withCoverage = ensureCoverageCandidates(normalized);
  return enforceRecommendationQuality(withCoverage);
}

function buildFallbackSmartArtItems(title, userPrompt) {
  const itemLines = cleanSentence(userPrompt, 320)
    .split(/\s*\|\s*|\s*;\s*|\r?\n+/)
    .map((line) => cleanSentence(line, 64))
    .filter((line) => line.length > 0);

  const selected = [];
  for (const line of itemLines) {
    if (selected.length >= 6) {
      break;
    }
    selected.push({ title: line });
  }

  if (selected.length >= 3) {
    return selected;
  }

  const safeTitle = cleanSentence(title || "Goal", 52) || "Goal";
  return [
    { title: `Define ${safeTitle}` },
    { title: "Prioritize key drivers" },
    { title: "Execute key actions" },
    { title: "Track outcomes" },
  ];
}

function fallbackPlan(selectedRecommendation, userPrompt) {
  const title = selectedRecommendation?.title || "Apply recommendation";
  const outputType = canonicalizeOutputType(selectedRecommendation?.outputType);
  const content = isSmartArtType(outputType)
    ? {
        smartArt: {
          layout: "process",
          title,
          items: buildFallbackSmartArtItems(title, userPrompt),
        },
      }
    : {
        text: `Apply recommendation: ${title}\nUser prompt: ${userPrompt}`,
      };
  return {
    planId: `plan-${Date.now()}`,
    summary: `Fallback plan for "${title}"`,
    requiresConfirmation: true,
    warnings: ["Generated fallback plan due to model parsing uncertainty."],
    operations: [
      {
        type: "insert",
        target: "content-placeholder",
        anchor: { strategy: "placeholder", ref: "body" },
        content,
        styleBindings: {
          font: "theme.body",
          color: "theme.text1",
          spacing: "theme.medium",
        },
        constraints: {
          avoidOverlap: true,
          keepEditable: true,
          preserveTheme: true,
        },
      },
    ],
  };
}

function normalizePlan(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const operations = Array.isArray(raw.operations) ? raw.operations : [];
  return {
    planId: typeof raw.planId === "string" && raw.planId.trim() ? raw.planId.trim() : `plan-${Date.now()}`,
    summary: typeof raw.summary === "string" ? raw.summary.trim() : "Execution plan",
    requiresConfirmation:
      typeof raw.requiresConfirmation === "boolean" ? raw.requiresConfirmation : true,
    warnings: Array.isArray(raw.warnings) ? raw.warnings.filter((v) => typeof v === "string") : [],
    operations: operations
      .map((op) => ({
        type: typeof op.type === "string" ? op.type : "update",
        target: typeof op.target === "string" ? op.target : "content-placeholder",
        anchor:
          op.anchor && typeof op.anchor === "object"
            ? {
                strategy:
                  typeof op.anchor.strategy === "string" ? op.anchor.strategy : "placeholder",
                ref: typeof op.anchor.ref === "string" ? op.anchor.ref : "body",
              }
            : { strategy: "placeholder", ref: "body" },
        content: op.content && typeof op.content === "object" ? op.content : {},
        styleBindings:
          op.styleBindings && typeof op.styleBindings === "object" ? op.styleBindings : {},
        constraints:
          op.constraints && typeof op.constraints === "object"
            ? op.constraints
            : { avoidOverlap: true, keepEditable: true, preserveTheme: true },
      }))
      .filter((op) => typeof op.type === "string"),
  };
}

function getImageContentObject(content) {
  if (!content || typeof content !== "object") {
    return null;
  }

  if (content.image && typeof content.image === "object") {
    return content.image;
  }

  if (typeof content.image === "string" && content.image.trim()) {
    return { url: content.image.trim(), alt: "" };
  }

  if (typeof content.imageUrl === "string" && content.imageUrl.trim()) {
    return { url: content.imageUrl.trim(), alt: "" };
  }

  return null;
}

async function enrichPlanImages(plan, { selectedRecommendation, userPrompt, slideContext }) {
  if (!plan || !Array.isArray(plan.operations) || plan.operations.length === 0) {
    return plan;
  }

  const warnings = Array.isArray(plan.warnings) ? [...plan.warnings] : [];
  const operations = [];

  for (const operation of plan.operations) {
    const op = operation && typeof operation === "object" ? { ...operation } : operation;
    const imageContent = getImageContentObject(op?.content);
    if (!imageContent) {
      operations.push(op);
      continue;
    }

    let resolved = null;
    try {
      resolved = await resolvePresentationImage({
        selectedRecommendation,
        userPrompt,
        slideContext,
        content: op.content,
      });
    } catch (_error) {
      resolved = null;
    }

    if (resolved && resolved.url) {
      const nextContent = op.content && typeof op.content === "object" ? { ...op.content } : {};
      const nextImage = {
        ...(typeof nextContent.image === "object" ? nextContent.image : {}),
        url: resolved.url,
        alt:
          (typeof nextContent.image?.alt === "string" && nextContent.image.alt.trim())
            ? nextContent.image.alt.trim()
            : cleanSentence(selectedRecommendation?.title || "Slide visual", 90),
        query: resolved.query || nextContent.image?.query || "",
      };
      nextContent.image = nextImage;
      delete nextContent.imageUrl;
      op.content = nextContent;
      operations.push(op);
      continue;
    }

    warnings.push(
      `Image for operation "${cleanSentence(op?.target || "content", 60)}" could not be validated from search providers.`
    );
    operations.push(op);
  }

  return {
    ...plan,
    operations,
    warnings: Array.from(new Set(warnings)),
  };
}

async function generateRecommendations({ userPrompt, slideContext }) {
  const contextSummary = summarizeContext(slideContext);
  const promptPayload = JSON.stringify({ userPrompt, contextSummary }, null, 2);
  debugLog("LLM recommendation prompt payload", promptPayload);

  const modelResponse = await generateStructuredJson({
    systemPrompt: buildRecommendationPrompt(),
    userPrompt: promptPayload,
    temperature: 0.3,
  });
  debugLog("LLM recommendation raw response", modelResponse);

  const parsed = tryParseJson(modelResponse);
  const normalized = normalizeRecommendations(parsed);
  const validation = validateRecommendations(normalized);
  if (normalized.length > 0 && validation.valid) {
    return normalized;
  }

  return fallbackRecommendations(userPrompt);
}

async function generateExecutionPlan({ selectedRecommendation, userPrompt, slideContext }) {
  const contextSummary = summarizeContext(slideContext);
  const promptPayload = JSON.stringify(
    { selectedRecommendation, userPrompt, contextSummary },
    null,
    2
  );
  debugLog("LLM execution-plan prompt payload", promptPayload);

  const modelResponse = await generateStructuredJson({
    systemPrompt: buildPlanPrompt(),
    userPrompt: promptPayload,
    temperature: 0.2,
  });
  debugLog("LLM execution-plan raw response", modelResponse);

  const parsed = tryParseJson(modelResponse);
  const normalized = normalizePlan(parsed);
  if (normalized && normalized.operations.length > 0) {
    const enriched = await enrichPlanImages(normalized, {
      selectedRecommendation,
      userPrompt,
      slideContext,
    });

    const validation = validateExecutionPlan(enriched);
    if (!validation.valid) {
      return fallbackPlan(selectedRecommendation, userPrompt);
    }

    if (validation.policyWarnings.length > 0) {
      const mergedWarnings = Array.from(
        new Set([...(enriched.warnings || []), ...validation.policyWarnings])
      );
      return {
        ...enriched,
        warnings: mergedWarnings,
        requiresConfirmation: true,
      };
    }

    return enriched;
  }

  return fallbackPlan(selectedRecommendation, userPrompt);
}

module.exports = {
  generateRecommendations,
  generateExecutionPlan,
};
