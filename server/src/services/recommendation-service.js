const { generateStructuredJson } = require("./ollama-client");
const { buildRecommendationPrompt, buildPlanPrompt } = require("./prompts");
const { tryParseJson } = require("../utils/json");
const { validateRecommendations, ALLOWED_OUTPUT_TYPES } = require("../validation/recommendations");
const { validateExecutionPlan } = require("../validation/execution-plan");

const DEBUG_LOGS = process.env.DEBUG_LOGS === "true" || process.env.NODE_ENV !== "production";

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
      id: "convert-to-table",
      title: "Convert to table",
      description: "Restructure content into a clean table format.",
      outputType: "table",
      confidence: 0.7,
      applyHints: ["use-content-placeholder", "avoid-overlap"],
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

function defaultHintsForType(outputType) {
  switch (outputType) {
    case "table":
      return ["prefer-table-shape", "preserve-theme", "avoid-overlap"];
    case "image":
      return ["insert-image-shape", "preserve-theme", "avoid-overlap"];
    case "chart":
      return ["prefer-chart-shape", "preserve-theme", "avoid-overlap"];
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
    case "summary":
      return "Condense slide content into concise key points for faster understanding.";
    case "layout-improvement":
      return "Improve alignment, spacing, and hierarchy to increase clarity and readability.";
    default:
      return "Apply a focused improvement that preserves design and keeps content editable.";
  }
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

  const byConfidence = deduped.sort((a, b) => b.confidence - a.confidence);

  const selected = [];
  const usedTypes = new Set();

  for (const item of byConfidence) {
    if (!usedTypes.has(item.outputType)) {
      selected.push(item);
      usedTypes.add(item.outputType);
    }
    if (selected.length >= 4) {
      break;
    }
  }

  if (selected.length < 3) {
    for (const item of byConfidence) {
      if (selected.indexOf(item) === -1) {
        selected.push(item);
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
      outputType: typeof item.outputType === "string" ? item.outputType : "other",
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
      applyHints:
        (item.applyHints.length > 0 ? item.applyHints : defaultHintsForType(item.outputType))
          .slice(0, 8)
          .map((hint) => cleanSentence(hint, 80)),
    }))
    .filter((item) => item.title.length > 0)
    .map((item) => ({
      ...item,
      id: toKebabId(item.id || item.title, "rec"),
      confidence: Math.max(0.35, Math.min(0.98, item.confidence)),
    }));

  return enforceRecommendationQuality(normalized);
}

function fallbackPlan(selectedRecommendation, userPrompt) {
  const title = selectedRecommendation?.title || "Apply recommendation";
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
        content: {
          text: `Apply recommendation: ${title}\nUser prompt: ${userPrompt}`,
        },
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
    return normalized.sort((a, b) => b.confidence - a.confidence);
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
    const validation = validateExecutionPlan(normalized);
    if (!validation.valid) {
      return fallbackPlan(selectedRecommendation, userPrompt);
    }

    if (validation.policyWarnings.length > 0) {
      const mergedWarnings = Array.from(
        new Set([...(normalized.warnings || []), ...validation.policyWarnings])
      );
      return {
        ...normalized,
        warnings: mergedWarnings,
        requiresConfirmation: true,
      };
    }

    return normalized;
  }

  return fallbackPlan(selectedRecommendation, userPrompt);
}

module.exports = {
  generateRecommendations,
  generateExecutionPlan,
};
