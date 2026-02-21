const ALLOWED_OUTPUT_TYPES = new Set([
  "list",
  "table",
  "chart",
  "image",
  "diagram",
  "summary",
  "layout-improvement",
  "other",
]);

function validateRecommendationItem(item) {
  const errors = [];
  if (!item || typeof item !== "object") {
    return { valid: false, errors: ["Recommendation item must be an object"] };
  }

  if (typeof item.id !== "string" || item.id.trim().length === 0) {
    errors.push("Recommendation id is required");
  }

  if (typeof item.title !== "string" || item.title.trim().length === 0) {
    errors.push(`Recommendation ${item.id || "<unknown>"} requires a title`);
  }

  if (typeof item.description !== "string") {
    errors.push(`Recommendation ${item.id || "<unknown>"} requires a description string`);
  }

  if (!ALLOWED_OUTPUT_TYPES.has(item.outputType)) {
    errors.push(`Recommendation ${item.id || "<unknown>"} has invalid outputType`);
  }

  if (typeof item.confidence !== "number" || item.confidence < 0 || item.confidence > 1) {
    errors.push(`Recommendation ${item.id || "<unknown>"} confidence must be between 0 and 1`);
  }

  if (!Array.isArray(item.applyHints)) {
    errors.push(`Recommendation ${item.id || "<unknown>"} applyHints must be an array`);
  }

  return { valid: errors.length === 0, errors };
}

function validateRecommendations(recommendations) {
  const errors = [];
  if (!Array.isArray(recommendations)) {
    return { valid: false, errors: ["recommendations must be an array"] };
  }

  if (recommendations.length < 1 || recommendations.length > 6) {
    errors.push("recommendations must contain between 1 and 6 items");
  }

  for (const item of recommendations) {
    const result = validateRecommendationItem(item);
    if (!result.valid) {
      errors.push(...result.errors);
    }
  }

  return { valid: errors.length === 0, errors };
}

module.exports = {
  validateRecommendations,
  ALLOWED_OUTPUT_TYPES,
};
