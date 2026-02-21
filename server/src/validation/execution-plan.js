const ALLOWED_OPERATION_TYPES = new Set(["insert", "update", "transform", "delete"]);
const ALLOWED_ANCHOR_STRATEGIES = new Set(["placeholder", "selection", "free-region"]);

function validateOperation(op, index) {
  const errors = [];
  const label = `operations[${index}]`;

  if (!op || typeof op !== "object") {
    return { valid: false, errors: [`${label} must be an object`] };
  }

  if (!ALLOWED_OPERATION_TYPES.has(op.type)) {
    errors.push(`${label}.type is invalid`);
  }

  if (typeof op.target !== "string" || op.target.trim().length === 0) {
    errors.push(`${label}.target is required`);
  }

  if (!op.anchor || typeof op.anchor !== "object") {
    errors.push(`${label}.anchor is required`);
  } else {
    if (!ALLOWED_ANCHOR_STRATEGIES.has(op.anchor.strategy)) {
      errors.push(`${label}.anchor.strategy is invalid`);
    }
    if (typeof op.anchor.ref !== "string" || op.anchor.ref.trim().length === 0) {
      errors.push(`${label}.anchor.ref is required`);
    }
  }

  if (!op.constraints || typeof op.constraints !== "object") {
    errors.push(`${label}.constraints is required`);
  }

  return { valid: errors.length === 0, errors };
}

function checkPolicyRisks(plan) {
  const warnings = [];
  const hasDelete = plan.operations.some((op) => op.type === "delete");
  if (hasDelete) {
    warnings.push("Plan contains delete operations and requires explicit confirmation.");
  }

  const hasFreeRegion = plan.operations.some((op) => op.anchor?.strategy === "free-region");
  if (hasFreeRegion) {
    warnings.push("Plan uses free-region anchoring; renderer must run collision checks before apply.");
  }

  return warnings;
}

function validateExecutionPlan(plan) {
  const errors = [];
  if (!plan || typeof plan !== "object") {
    return { valid: false, errors: ["plan must be an object"], policyWarnings: [] };
  }

  if (typeof plan.planId !== "string" || plan.planId.trim().length === 0) {
    errors.push("plan.planId is required");
  }

  if (typeof plan.summary !== "string" || plan.summary.trim().length === 0) {
    errors.push("plan.summary is required");
  }

  if (!Array.isArray(plan.operations) || plan.operations.length < 1 || plan.operations.length > 15) {
    errors.push("plan.operations must contain between 1 and 15 items");
  } else {
    for (let i = 0; i < plan.operations.length; i += 1) {
      const result = validateOperation(plan.operations[i], i);
      if (!result.valid) {
        errors.push(...result.errors);
      }
    }
  }

  const policyWarnings = errors.length === 0 ? checkPolicyRisks(plan) : [];
  return {
    valid: errors.length === 0,
    errors,
    policyWarnings,
  };
}

module.exports = {
  validateExecutionPlan,
};

