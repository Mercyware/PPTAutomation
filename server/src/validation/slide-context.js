function validateSlideContext(slideContext) {
  const errors = [];

  if (!slideContext || typeof slideContext !== "object") {
    return { valid: false, errors: ["slideContext must be an object"] };
  }

  const objects = Array.isArray(slideContext.objects) ? slideContext.objects : [];
  if (objects.length > 250) {
    errors.push("slideContext.objects exceeds max length of 250");
  }

  const invalidObject = objects.find((obj) => !obj || typeof obj !== "object" || typeof obj.id !== "string");
  if (invalidObject) {
    errors.push("Every slideContext object must contain an id string");
  }

  if (slideContext.slide && typeof slideContext.slide !== "object") {
    errors.push("slideContext.slide must be an object when provided");
  }

  if (slideContext.selection && typeof slideContext.selection !== "object") {
    errors.push("slideContext.selection must be an object when provided");
  }

  return { valid: errors.length === 0, errors };
}

module.exports = {
  validateSlideContext,
};

