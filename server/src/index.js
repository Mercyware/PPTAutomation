const express = require("express");
const cors = require("cors");
const dotenv = require("dotenv");

dotenv.config();

const { generateRecommendations, generateExecutionPlan } = require("./services/recommendation-service");
const { findReferenceForItem } = require("./services/reference-service");
const { ACTIVE_PROVIDER, ACTIVE_MODEL } = require("./services/ollama-client");
const { validateSlideContext } = require("./validation/slide-context");

const DEBUG_LOGS = process.env.DEBUG_LOGS === "true" || process.env.NODE_ENV !== "production";

function debugLog(label, payload) {
  if (!DEBUG_LOGS) {
    return;
  }
  const timestamp = new Date().toISOString();
  console.log(`[server][${timestamp}] ${label}`);
  if (payload !== undefined) {
    try {
      console.log(typeof payload === "string" ? payload : JSON.stringify(payload, null, 2));
    } catch (_error) {
      console.log(String(payload));
    }
  }
}

const app = express();
const port = Number(process.env.PORT);
if (!Number.isFinite(port) || port <= 0) {
  throw new Error("Invalid or missing PORT in .env");
}

app.use(cors());
app.use(express.json({ limit: "2mb" }));

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    service: "ppt-automation-server",
    provider: ACTIVE_PROVIDER,
    model: ACTIVE_MODEL,
  });
});

app.post("/api/recommendations", async (req, res) => {
  try {
    const { userPrompt, slideContext } = req.body || {};
    debugLog("Incoming /api/recommendations", {
      userPrompt,
      objectCount: Array.isArray(slideContext?.objects) ? slideContext.objects.length : 0,
      selectionCount: Array.isArray(slideContext?.selection?.shapeIds)
        ? slideContext.selection.shapeIds.length
        : 0,
    });
    if (!userPrompt || typeof userPrompt !== "string") {
      return res.status(400).json({ error: "userPrompt is required and must be a string" });
    }

    const contextValidation = validateSlideContext(slideContext || {});
    if (!contextValidation.valid) {
      return res.status(400).json({
        error: "Invalid slideContext",
        details: contextValidation.errors,
      });
    }

    const recommendations = await generateRecommendations({
      userPrompt,
      slideContext: slideContext || {},
    });

    return res.json({ recommendations });
  } catch (error) {
    return res.status(500).json({
      error: "Failed to generate recommendations",
      details: error.message,
    });
  }
});

app.post("/api/plans", async (req, res) => {
  try {
    const { selectedRecommendation, userPrompt, slideContext } = req.body || {};
    debugLog("Incoming /api/plans", {
      userPrompt,
      selectedRecommendation: selectedRecommendation
        ? {
            id: selectedRecommendation.id,
            title: selectedRecommendation.title,
            outputType: selectedRecommendation.outputType,
          }
        : null,
      objectCount: Array.isArray(slideContext?.objects) ? slideContext.objects.length : 0,
    });
    if (!selectedRecommendation || typeof selectedRecommendation !== "object") {
      return res.status(400).json({ error: "selectedRecommendation is required" });
    }

    const contextValidation = validateSlideContext(slideContext || {});
    if (!contextValidation.valid) {
      return res.status(400).json({
        error: "Invalid slideContext",
        details: contextValidation.errors,
      });
    }

    const plan = await generateExecutionPlan({
      selectedRecommendation,
      userPrompt: typeof userPrompt === "string" ? userPrompt : "",
      slideContext: slideContext || {},
    });

    return res.json({ plan });
  } catch (error) {
    return res.status(500).json({
      error: "Failed to generate execution plan",
      details: error.message,
    });
  }
});

app.post("/api/references", async (req, res) => {
  try {
    const { itemText, slideContext } = req.body || {};
    if (!itemText || typeof itemText !== "string") {
      return res.status(400).json({ error: "itemText is required and must be a string" });
    }

    debugLog("Incoming /api/references", {
      itemText: itemText.slice(0, 220),
      objectCount: Array.isArray(slideContext?.objects) ? slideContext.objects.length : 0,
    });

    const result = await findReferenceForItem({
      itemText,
      slideContext: slideContext || {},
    });
    return res.json(result);
  } catch (error) {
    return res.status(500).json({
      error: "Failed to find reference",
      details: error && error.message ? error.message : String(error),
    });
  }
});

app.listen(port, () => {
  console.log(`ppt-automation-server running on http://localhost:${port}`);
});
