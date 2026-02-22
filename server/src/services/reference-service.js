const { generateStructuredJson } = require("./ollama-client");
const { buildReferenceQueryPrompt, buildReferenceSelectionPrompt } = require("./prompts");
const { tryParseJson } = require("../utils/json");

const DEBUG_LOGS = process.env.DEBUG_LOGS === "true" || process.env.NODE_ENV !== "production";
const WIKIDATA_API = "https://www.wikidata.org/w/api.php";

function debugLog(label, payload) {
  if (!DEBUG_LOGS) {
    return;
  }
  const timestamp = new Date().toISOString();
  console.log(`[reference-service][${timestamp}] ${label}`);
  if (payload !== undefined) {
    try {
      console.log(typeof payload === "string" ? payload : JSON.stringify(payload, null, 2));
    } catch (_error) {
      console.log(String(payload));
    }
  }
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function safeSlice(value, maxLen) {
  return normalizeText(value).slice(0, maxLen);
}

function toTokens(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(Boolean);
}

function computeSimilarity(a, b) {
  const aTokens = new Set(toTokens(a));
  const bTokens = new Set(toTokens(b));
  if (!aTokens.size || !bTokens.size) {
    return 0;
  }
  let overlap = 0;
  for (const token of aTokens) {
    if (bTokens.has(token)) {
      overlap += 1;
    }
  }
  return overlap / Math.max(aTokens.size, bTokens.size);
}

function stripLowSignalTerms(tokens) {
  const stopwords = new Set([
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "globally",
    "has", "have", "in", "is", "it", "its", "of", "on", "or", "rank", "ranks",
    "ranking", "that", "the", "to", "was", "were", "with", "within", "world", "global",
  ]);
  return tokens.filter((token) => token.length > 2 && !stopwords.has(token) && !/^\d+$/.test(token));
}

function buildHeuristicQueries(itemText, slideContext) {
  const base = safeSlice(itemText, 220);
  const baseTokens = stripLowSignalTerms(toTokens(base));
  const queries = [];
  if (base) {
    queries.push(base);
  }
  if (baseTokens.length) {
    queries.push(baseTokens.slice(0, 8).join(" "));
  }

  const objects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const contextTerms = [];
  for (let i = 0; i < objects.length && contextTerms.length < 3; i += 1) {
    const text = normalizeText(objects[i]?.text || "");
    if (!text) {
      continue;
    }
    contextTerms.push(stripLowSignalTerms(toTokens(text)).slice(0, 5).join(" "));
  }
  for (const term of contextTerms) {
    if (term) {
      queries.push(`${baseTokens.slice(0, 5).join(" ")} ${term}`.trim());
    }
  }

  const tokenSet = new Set(baseTokens);
  if (tokenSet.has("gender") && tokenSet.has("equality")) {
    queries.push("global gender gap report");
  }
  if (tokenSet.has("inequality") && tokenSet.has("gender")) {
    queries.push("gender inequality index");
  }
  if (tokenSet.has("corruption")) {
    queries.push("corruption perceptions index");
  }
  if (tokenSet.has("development") && (tokenSet.has("human") || tokenSet.has("hdi"))) {
    queries.push("human development index");
  }
  if (tokenSet.has("competitiveness")) {
    queries.push("global competitiveness report");
  }

  return dedupeStrings(queries).slice(0, 4);
}

function dedupeStrings(values) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const normalized = normalizeText(value);
    if (!normalized) {
      continue;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    output.push(normalized);
  }
  return output;
}

async function buildAiQueries(itemText, slideContext) {
  const contextObjects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const contextSnippets = [];
  for (let i = 0; i < contextObjects.length && contextSnippets.length < 4; i += 1) {
    const text = safeSlice(contextObjects[i]?.text || "", 120);
    if (text) {
      contextSnippets.push(text);
    }
  }

  const payload = JSON.stringify({
    selectedItemText: safeSlice(itemText, 220),
    slideSnippets: contextSnippets,
  });

  try {
    const response = await generateStructuredJson({
      systemPrompt: buildReferenceQueryPrompt(),
      userPrompt: payload,
      temperature: 0.2,
    });
    const parsed = tryParseJson(response);
    const queries = Array.isArray(parsed?.queries) ? parsed.queries : [];
    return dedupeStrings(queries).slice(0, 4);
  } catch (error) {
    debugLog("AI query generation failed", error && error.message ? error.message : error);
    return [];
  }
}

function toSearchUrl(query) {
  const params = new URLSearchParams({
    action: "wbsearchentities",
    search: query,
    language: "en",
    limit: "8",
    format: "json",
  });
  return `${WIKIDATA_API}?${params.toString()}`;
}

async function fetchJsonWithTimeout(url, options) {
  const timeoutMs = Number(options?.timeoutMs || 8000);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      ...(options || {}),
      signal: controller.signal,
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status} for ${url}`);
    }
    return response.json();
  } finally {
    clearTimeout(timeout);
  }
}

async function searchWikidata(query) {
  const url = toSearchUrl(query);
  const payload = await fetchJsonWithTimeout(url, { timeoutMs: 10000 });
  return Array.isArray(payload?.search) ? payload.search : [];
}

async function fetchEntityDetails(entityIds) {
  if (!Array.isArray(entityIds) || !entityIds.length) {
    return {};
  }

  const params = new URLSearchParams({
    action: "wbgetentities",
    ids: entityIds.join("|"),
    props: "labels|descriptions|sitelinks|claims",
    languages: "en",
    format: "json",
  });
  const url = `${WIKIDATA_API}?${params.toString()}`;
  const payload = await fetchJsonWithTimeout(url, { timeoutMs: 10000 });
  return payload?.entities && typeof payload.entities === "object" ? payload.entities : {};
}

function extractOfficialWebsites(entity) {
  const claims = entity?.claims;
  const raw = Array.isArray(claims?.P856) ? claims.P856 : [];
  const urls = [];
  for (const claim of raw) {
    const value = claim?.mainsnak?.datavalue?.value;
    const normalized = normalizeHttpUrl(value);
    if (normalized) {
      urls.push(normalized);
    }
  }
  return dedupeStrings(urls);
}

function normalizeHttpUrl(value) {
  const raw = normalizeText(value);
  if (!raw) {
    return null;
  }
  try {
    const parsed = new URL(raw);
    if (!/^https?:$/i.test(parsed.protocol)) {
      return null;
    }
    return parsed.toString();
  } catch (_error) {
    return null;
  }
}

function normalizeEntityLabel(entity) {
  const label = entity?.labels?.en?.value || entity?.label || "";
  const desc = entity?.descriptions?.en?.value || entity?.description || "";
  return {
    label: safeSlice(label, 180),
    description: safeSlice(desc, 320),
  };
}

function confidenceBoostFromDomain(url) {
  try {
    const hostname = new URL(url).hostname.toLowerCase();
    if (
      hostname.endsWith(".gov") ||
      hostname.endsWith(".edu") ||
      hostname.endsWith(".int") ||
      hostname.endsWith(".org")
    ) {
      return 0.15;
    }
    if (hostname.includes("wikipedia.org")) {
      return 0.1;
    }
  } catch (_error) {
    return 0;
  }
  return 0.03;
}

function buildCandidatesForEntity(entity, searchQuery) {
  const { label, description } = normalizeEntityLabel(entity);
  const topicText = `${label} ${description}`.trim();
  const similarity = computeSimilarity(searchQuery, topicText);
  const candidates = [];
  const officialWebsites = extractOfficialWebsites(entity);
  for (const url of officialWebsites) {
    const score = Math.min(0.97, 0.5 + similarity * 0.35 + confidenceBoostFromDomain(url));
    candidates.push({
      title: label || "Official source",
      url,
      sourceType: "official-website",
      confidence: Number(score.toFixed(2)),
      reason: description || "Official website listed on Wikidata",
      wikidataId: entity?.id || null,
      query: searchQuery,
    });
  }

  const wikiTitle = entity?.sitelinks?.enwiki?.title;
  if (typeof wikiTitle === "string" && wikiTitle.trim()) {
    const wikiUrl = `https://en.wikipedia.org/wiki/${encodeURIComponent(wikiTitle.replace(/\s+/g, "_"))}`;
    const score = Math.min(0.9, 0.44 + similarity * 0.32 + confidenceBoostFromDomain(wikiUrl));
    candidates.push({
      title: label || wikiTitle,
      url: wikiUrl,
      sourceType: "wikipedia",
      confidence: Number(score.toFixed(2)),
      reason: description || "Relevant encyclopedia article",
      wikidataId: entity?.id || null,
      query: searchQuery,
    });
  }

  return candidates;
}

async function searchSemanticScholar(query) {
  const params = new URLSearchParams({
    query,
    limit: "6",
    fields: "title,url,year,externalIds,venue",
  });
  const url = `https://api.semanticscholar.org/graph/v1/paper/search?${params.toString()}`;
  const payload = await fetchJsonWithTimeout(url, { timeoutMs: 12000 });
  return Array.isArray(payload?.data) ? payload.data : [];
}

function buildCandidatesForSemanticResult(item, searchQuery) {
  const title = safeSlice(item?.title || "", 200);
  if (!title) {
    return [];
  }

  const primaryUrl = normalizeHttpUrl(item?.url);
  const doiValue = typeof item?.externalIds?.DOI === "string" ? item.externalIds.DOI.trim() : "";
  const doiUrl = doiValue ? normalizeHttpUrl(`https://doi.org/${doiValue}`) : null;
  const urls = dedupeStrings([doiUrl, primaryUrl]);
  const year = Number(item?.year || 0);
  const venue = safeSlice(item?.venue || "", 90);
  const similarity = computeSimilarity(searchQuery, title);
  const reason =
    `Research match${year > 0 ? ` (${year})` : ""}${venue ? ` in ${venue}` : ""}`.trim();
  const candidates = [];
  for (const url of urls) {
    const score = Math.min(0.74, 0.32 + similarity * 0.28 + confidenceBoostFromDomain(url) * 0.45);
    candidates.push({
      title,
      url,
      sourceType: "research-paper",
      confidence: Number(score.toFixed(2)),
      reason,
      wikidataId: null,
      query: searchQuery,
    });
  }
  return candidates;
}

function dedupeCandidates(candidates) {
  const seen = new Set();
  const output = [];
  for (const candidate of candidates) {
    const url = normalizeHttpUrl(candidate?.url);
    if (!url) {
      continue;
    }
    const key = url.toLowerCase().replace(/\/+$/, "");
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    output.push({
      ...candidate,
      url,
    });
  }
  return output;
}

async function isUrlReachable(url) {
  const normalized = normalizeHttpUrl(url);
  if (!normalized) {
    return false;
  }

  const timeoutMs = 8000;
  const headController = new AbortController();
  const headTimeout = setTimeout(() => headController.abort(), timeoutMs);
  try {
    const head = await fetch(normalized, {
      method: "HEAD",
      redirect: "follow",
      signal: headController.signal,
    });
    if (head.ok) {
      return true;
    }
    if (head.status >= 400 && head.status !== 405) {
      return false;
    }
  } catch (_error) {
    // fallback to GET
  } finally {
    clearTimeout(headTimeout);
  }

  const getController = new AbortController();
  const getTimeout = setTimeout(() => getController.abort(), timeoutMs);
  try {
    const get = await fetch(normalized, {
      method: "GET",
      redirect: "follow",
      signal: getController.signal,
    });
    return get.status >= 200 && get.status < 400;
  } catch (_error) {
    return false;
  } finally {
    clearTimeout(getTimeout);
  }
}

async function validateTopCandidates(candidates, limit) {
  const copy = Array.isArray(candidates) ? [...candidates] : [];
  const max = Math.max(0, Number(limit || 8));
  for (let i = 0; i < copy.length && i < max; i += 1) {
    const reachable = await isUrlReachable(copy[i].url);
    copy[i] = {
      ...copy[i],
      reachable,
      confidence: Number(Math.min(0.99, copy[i].confidence + (reachable ? 0.08 : -0.18)).toFixed(2)),
    };
  }
  return copy;
}

function toSortedCandidates(candidates) {
  return [...candidates].sort((a, b) => {
    if (b.confidence !== a.confidence) {
      return b.confidence - a.confidence;
    }
    return String(a.title || "").localeCompare(String(b.title || ""));
  });
}

function clampString(value, maxLen) {
  return safeSlice(value, maxLen);
}

function clampConfidence(value, fallback) {
  const n = Number(value);
  if (!Number.isFinite(n)) {
    return fallback;
  }
  return Math.max(0, Math.min(1, n));
}

async function buildLlmReferenceCandidate(itemText, slideContext) {
  const contextObjects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const snippets = [];
  for (let i = 0; i < contextObjects.length && snippets.length < 6; i += 1) {
    const text = safeSlice(contextObjects[i]?.text || "", 140);
    if (text) {
      snippets.push(text);
    }
  }

  const payload = JSON.stringify({
    selectedItemText: safeSlice(itemText, 260),
    slideSnippets: snippets,
  });

  try {
    const response = await generateStructuredJson({
      systemPrompt: buildReferenceSelectionPrompt(),
      userPrompt: payload,
      temperature: 0.1,
    });
    const parsed = tryParseJson(response);
    const reference = parsed?.reference && typeof parsed.reference === "object" ? parsed.reference : null;
    if (!reference) {
      return null;
    }

    const url = normalizeHttpUrl(reference.url);
    if (!url) {
      return null;
    }

    let reachable = false;
    try {
      reachable = await isUrlReachable(url);
    } catch (_error) {
      reachable = false;
    }

    return {
      reference: {
        title: clampString(reference.title, 180) || "Reference",
        url,
        reason: clampString(reference.reason, 220) || "LLM-selected reference",
        confidence: Number(clampConfidence(reference.confidence, reachable ? 0.72 : 0.62).toFixed(2)),
        sourceType: "llm-direct",
        query: clampString(itemText, 260),
        reachable: Boolean(reachable),
      },
      alternatives: [],
    };
  } catch (error) {
    debugLog("LLM direct reference failed", error && error.message ? error.message : error);
    return null;
  }
}

function buildSearchFallbackReference(itemText) {
  const query = clampString(itemText, 220) || "reference";
  const url = `https://en.wikipedia.org/w/index.php?search=${encodeURIComponent(query)}`;
  return {
    reference: {
      title: clampString(`Wikipedia search: ${query}`, 180),
      url,
      reason: "Fallback reference search link",
      confidence: 0.4,
      sourceType: "fallback-search",
      query,
      reachable: true,
    },
    alternatives: [],
  };
}

function buildTopicMappedReference(itemText) {
  const text = normalizeText(itemText).toLowerCase();
  if (!text) {
    return null;
  }

  if (text.includes("gender") && (text.includes("equality") || text.includes("gap"))) {
    return {
      reference: {
        title: "Global Gender Gap Report",
        url: "https://en.wikipedia.org/wiki/Global_Gender_Gap_Report",
        reason: "Stable reference page for the gender-gap ranking framework.",
        confidence: 0.7,
        sourceType: "topic-map",
        query: clampString(itemText, 220),
        reachable: true,
      },
      alternatives: [],
    };
  }

  if (text.includes("corruption") && text.includes("index")) {
    return {
      reference: {
        title: "Corruption Perceptions Index",
        url: "https://en.wikipedia.org/wiki/Corruption_Perceptions_Index",
        reason: "Stable reference page for CPI ranking methodology.",
        confidence: 0.7,
        sourceType: "topic-map",
        query: clampString(itemText, 220),
        reachable: true,
      },
      alternatives: [],
    };
  }

  if ((text.includes("human development") || text.includes("hdi")) && text.includes("index")) {
    return {
      reference: {
        title: "Human Development Index",
        url: "https://en.wikipedia.org/wiki/Human_Development_Index",
        reason: "Stable reference page for HDI ranking methodology.",
        confidence: 0.7,
        sourceType: "topic-map",
        query: clampString(itemText, 220),
        reachable: true,
      },
      alternatives: [],
    };
  }

  return null;
}

async function findReferenceForItem({ itemText, slideContext }) {
  const normalizedText = clampString(itemText, 300);
  if (!normalizedText) {
    throw new Error("itemText is required for reference lookup");
  }

  const llmDirect = await buildLlmReferenceCandidate(normalizedText, slideContext);
  if (llmDirect && llmDirect.reference && llmDirect.reference.url) {
    return llmDirect;
  }

  const mapped = buildTopicMappedReference(normalizedText);
  if (mapped && mapped.reference && mapped.reference.url) {
    return mapped;
  }

  const heuristicQueries = buildHeuristicQueries(normalizedText, slideContext);
  const aiQueries = await buildAiQueries(normalizedText, slideContext);
  const queries = dedupeStrings([...aiQueries, ...heuristicQueries]).slice(0, 6);
  if (!queries.length) {
    return buildSearchFallbackReference(normalizedText);
  }

  debugLog("Reference search queries", queries);
  const entityById = new Map();
  for (const query of queries) {
    try {
      const searchResults = await searchWikidata(query);
      for (const row of searchResults) {
        const entityId = String(row?.id || "").trim();
        if (!entityId || entityById.has(entityId)) {
          continue;
        }
        entityById.set(entityId, row);
      }
    } catch (error) {
      debugLog("Wikidata search failed", {
        query,
        error: error && error.message ? error.message : String(error),
      });
    }
  }
  const allCandidates = [];

  if (entityById.size > 0) {
    const entityIds = Array.from(entityById.keys()).slice(0, 28);
    const details = await fetchEntityDetails(entityIds);
    for (const id of entityIds) {
      const row = entityById.get(id) || {};
      const entity = details[id] && typeof details[id] === "object" ? details[id] : row;
      if (!entity || typeof entity !== "object") {
        continue;
      }

      entity.id = entity.id || id;
      const queryAnchor = queries[0];
      const candidates = buildCandidatesForEntity(entity, queryAnchor);
      allCandidates.push(...candidates);
    }
  }

  if (allCandidates.length === 0) {
    const semanticQueries = queries.slice(0, 2);
    for (const query of semanticQueries) {
      try {
        const papers = await searchSemanticScholar(query);
        for (const paper of papers) {
          allCandidates.push(...buildCandidatesForSemanticResult(paper, query));
        }
      } catch (error) {
        debugLog("Semantic Scholar search failed", {
          query,
          error: error && error.message ? error.message : String(error),
        });
      }
    }
  }

  const dedupedCandidates = dedupeCandidates(allCandidates);
  if (!dedupedCandidates.length) {
    return buildSearchFallbackReference(normalizedText);
  }

  const rankedCandidates = toSortedCandidates(dedupedCandidates);
  const validated = await validateTopCandidates(rankedCandidates, 12);
  const sorted = toSortedCandidates(validated);
  const best = sorted.find((candidate) => candidate.reachable);
  if (!best) {
    return buildSearchFallbackReference(normalizedText);
  }

  return {
    reference: {
      title: clampString(best.title, 180) || "Reference",
      url: best.url,
      reason: clampString(best.reason, 220) || "Best available match from reference search",
      confidence: best.confidence,
      sourceType: best.sourceType,
      query: best.query,
      reachable: Boolean(best.reachable),
    },
    alternatives: sorted.slice(1, 4).map((item) => ({
      title: clampString(item.title, 180),
      url: item.url,
      confidence: item.confidence,
      sourceType: item.sourceType,
      reachable: Boolean(item.reachable),
    })),
  };
}

module.exports = {
  findReferenceForItem,
};
