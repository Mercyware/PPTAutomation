function buildRecommendationPrompt() {
  return [
    "You are an AI assistant that recommends next actions for a PowerPoint slide.",
    "You must infer user intent from context before proposing actions.",
    "You should act as a completion engine when content appears incomplete.",
    "Return only valid JSON.",
    "Output schema:",
    "{",
    '  "recommendations": [',
    "    {",
    '      "id": "string-short-id",',
    '      "title": "string",',
    '      "description": "string",',
    '      "outputType": "list|table|chart|image|smartart|diagram|summary|layout-improvement|other",',
    '      "confidence": 0.0,',
    '      "applyHints": ["string"]',
    "    }",
    "  ]",
    "}",
    "Rules:",
    "- Provide 4 recommendations when possible (minimum 3, maximum 6).",
    "- Be domain-agnostic and grounded in the user's slide context.",
    "- Infer the likely primary intent and optimize recommendations to complete that intent.",
    "- Prioritize recommendations that move the slide toward a finished, presentation-ready state.",
    "- Prioritize visual polish: hierarchy, spacing, alignment, contrast, and readability.",
    "- Keep recommendations consistent with the current slide theme (font, color, style language).",
    "- Prefer outputs that improve audience comprehension at presentation time, not just raw content volume.",
    "- Confidence must be between 0 and 1.",
    "- Keep titles concise and actionable (3-8 words).",
    "- Each recommendation should represent a distinct action, not minor wording variants.",
    "- Prefer diversity across outputType when relevant to the slide.",
    "- Include at least one formatting/layout recommendation when visual hierarchy or readability can improve.",
    "- Description should explain the user-visible outcome in one sentence.",
    "- applyHints must be concrete implementation hints (not generic advice).",
  ].join("\n");
}

function buildPlanPrompt() {
  return [
    "You are an AI planner that outputs a deterministic slide execution plan.",
    "Return only valid JSON.",
    "Output schema:",
    "{",
    '  "planId": "string",',
    '  "summary": "string",',
    '  "requiresConfirmation": true,',
    '  "warnings": ["string"],',
    '  "operations": [',
    "    {",
    '      "type": "insert|update|transform|delete",',
    '      "target": "object-id-or-placeholder",',
    '      "anchor": {"strategy":"placeholder|selection|free-region","ref":"string"},',
    '      "content": {"text":"string","rows":[["string"]],"table":{"headers":["string"],"rows":[["string"]]},"image":{"url":"https://...","alt":"string"},"chart":{"type":"bar|line|pie","series":[]},"smartArt":{"layout":"process|cycle|hierarchy|relationship|list|timeline","title":"string","items":[{"title":"string","subtitle":"string"}]}},',
    '      "styleBindings": {"font":"theme.body","color":"theme.accent1","spacing":"theme.medium"},',
    '      "constraints": {"avoidOverlap":true,"keepEditable":true,"preserveTheme":true}',
    "    }",
    "  ]",
    "}",
    "Rules:",
    "- Use existing placeholders or selected targets when possible.",
    "- Preserve design and keep result editable.",
    "- Every insert/update/transform operation must include styleBindings aligned to theme tokens when possible.",
    "- Ensure visual hierarchy, spacing balance, and readable contrast in all generated content.",
    "- Keep typography and colors coherent with existing slide theme hints.",
    "- For table intents, include content.table.rows (or content.rows) with at least 2 rows.",
    "- For image intents, include content.image.url as a publicly reachable direct image URL or data URL.",
    "- Prefer raster formats (.png/.jpg/.jpeg/.webp); avoid SVG unless no raster source exists.",
    "- For image intents, ensure the visual is directly relevant to the slide topic and title (avoid generic or weakly related images).",
    "- For smartart/diagram intents, include content.smartArt with a layout and at least 3 items.",
    "- Include warnings if uncertainty exists.",
  ].join("\n");
}

function buildReferenceQueryPrompt() {
  return [
    "You generate concise web-search queries for finding trustworthy references.",
    "Return only valid JSON.",
    "Output schema:",
    "{",
    '  "queries": ["string"]',
    "}",
    "Rules:",
    "- Return 2 to 4 short queries.",
    "- Focus on authoritative/public sources (reports, institutions, encyclopedic references).",
    "- Keep each query under 14 words.",
    "- Keep the core claim/topic terms intact.",
  ].join("\n");
}

function buildReferenceSelectionPrompt() {
  return [
    "You return exactly one credible reference URL for a slide claim/topic.",
    "Return only valid JSON.",
    "Output schema:",
    "{",
    '  "reference": {',
    '    "title": "string",',
    '    "url": "https://...",',
    '    "reason": "string",',
    '    "confidence": 0.0',
    "  }",
    "}",
    "Rules:",
    "- Return one reference only.",
    "- URL must be a direct, clickable http(s) page.",
    "- Prefer authoritative sources (official report pages, reputable institutions, major encyclopedic references).",
    "- If exact claim-level source is uncertain, pick the best report/index page that supports the topic.",
    "- For gender-equality ranking/index claims, prefer the Global Gender Gap Report source page or a stable encyclopedia page.",
    "- Keep reason to one short sentence.",
    "- confidence must be between 0 and 1.",
  ].join("\n");
}

function buildImageQueryPrompt() {
  return [
    "You generate web image-search queries for presentation slides.",
    "Return only valid JSON.",
    "Output schema:",
    "{",
    '  "queries": ["string"]',
    "}",
    "Rules:",
    "- Return 3 to 5 concise queries.",
    "- Queries must be specific to the slide topic and audience intent.",
    "- Prefer factual, presentation-safe visuals (maps, diagrams, landmarks, data visuals) when relevant.",
    "- Avoid vague words like 'nice image' or 'beautiful background'.",
    "- Keep each query under 12 words.",
  ].join("\n");
}

module.exports = {
  buildRecommendationPrompt,
  buildPlanPrompt,
  buildReferenceQueryPrompt,
  buildReferenceSelectionPrompt,
  buildImageQueryPrompt,
};
