const { generateStructuredJson } = require("./ollama-client");
const { buildImageQueryPrompt } = require("./prompts");
const { tryParseJson } = require("../utils/json");

const GOOGLE_CSE_API_KEY = process.env.GOOGLE_CSE_API_KEY;
const GOOGLE_CSE_CX = process.env.GOOGLE_CSE_CX;
const IMAGE_SEARCH_TIMEOUT_MS = Number(process.env.IMAGE_SEARCH_TIMEOUT_MS || 12000);

const DEBUG_LOGS = process.env.DEBUG_LOGS === "true" || process.env.NODE_ENV !== "production";

function debugLog(label, payload) {
  if (!DEBUG_LOGS) {
    return;
  }
  const timestamp = new Date().toISOString();
  console.log(`[image-search-service][${timestamp}] ${label}`);
  if (payload !== undefined) {
    try {
      console.log(typeof payload === "string" ? payload : JSON.stringify(payload, null, 2));
    } catch (_error) {
      console.log(String(payload));
    }
  }
}

function cleanText(value, maxLen) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, maxLen || 160);
}

function normalizeQuery(value) {
  return cleanText(value, 96)
    .replace(/[^\w\s-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizeHttpUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  try {
    const parsed = new URL(raw);
    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
      return "";
    }
    return parsed.toString();
  } catch (_error) {
    return "";
  }
}

function isSvgTypeOrUrl(contentType, url) {
  const mime = String(contentType || "").toLowerCase();
  if (mime.includes("image/svg+xml")) {
    return true;
  }
  return /\.svg(?:$|[?#])/i.test(String(url || ""));
}

function toImagePayload(content) {
  if (!content || typeof content !== "object") {
    return { url: "", alt: "", query: "" };
  }

  if (typeof content.image === "string" && content.image.trim()) {
    return { url: content.image.trim(), alt: "", query: "" };
  }

  if (content.image && typeof content.image === "object") {
    const image = content.image;
    return {
      url: cleanText(image.url || image.src || image.dataUrl, 400),
      alt: cleanText(image.alt || "", 120),
      query: cleanText(image.query || image.searchQuery || image.topic || image.keyword || "", 120),
    };
  }

  return {
    url: cleanText(content.imageUrl || "", 400),
    alt: "",
    query: cleanText(content.imageQuery || "", 120),
  };
}

function deriveFallbackQueries({ selectedRecommendation, userPrompt, slideContext, imagePayload }) {
  const queries = [];
  const seen = new Set();
  const push = (value) => {
    const normalized = normalizeQuery(value);
    if (!normalized) return;
    const key = normalized.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    queries.push(normalized);
  };

  const title = cleanText(selectedRecommendation?.title, 90);
  const description = cleanText(selectedRecommendation?.description, 140);
  const prompt = cleanText(userPrompt, 220);
  const slideObjects = Array.isArray(slideContext?.objects) ? slideContext.objects : [];
  const slideText = slideObjects
    .map((obj) => cleanText(obj?.text, 80))
    .filter(Boolean)
    .slice(0, 5)
    .join(" ");

  push(imagePayload?.query);
  push(imagePayload?.alt);
  push(title);
  push(description);
  push(prompt);
  push(slideText);

  return queries.slice(0, 6);
}

async function generateImageQueries({ selectedRecommendation, userPrompt, slideContext, imagePayload }) {
  const fallback = deriveFallbackQueries({ selectedRecommendation, userPrompt, slideContext, imagePayload });
  const promptPayload = JSON.stringify(
    {
      recommendation: {
        title: selectedRecommendation?.title || "",
        description: selectedRecommendation?.description || "",
        outputType: selectedRecommendation?.outputType || "",
      },
      userPrompt: cleanText(userPrompt, 360),
      imageHint: {
        alt: imagePayload?.alt || "",
        query: imagePayload?.query || "",
      },
      slideKeywords: fallback.slice(0, 4),
    },
    null,
    2
  );

  try {
    const response = await generateStructuredJson({
      systemPrompt: buildImageQueryPrompt(),
      userPrompt: promptPayload,
      temperature: 0.2,
    });
    const parsed = tryParseJson(response);
    const llmQueries = Array.isArray(parsed?.queries) ? parsed.queries : [];
    const all = [...llmQueries, ...fallback];
    const deduped = [];
    const seen = new Set();
    for (const query of all) {
      const normalized = normalizeQuery(query);
      if (!normalized) continue;
      const key = normalized.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      deduped.push(normalized);
      if (deduped.length >= 8) break;
    }
    return deduped.length ? deduped : fallback;
  } catch (_error) {
    return fallback;
  }
}

async function fetchJsonWithTimeout(url) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), IMAGE_SEARCH_TIMEOUT_MS);
  try {
    const response = await fetch(url, {
      method: "GET",
      redirect: "follow",
      signal: controller.signal,
    });
    if (!response.ok) {
      return null;
    }
    return await response.json();
  } catch (_error) {
    return null;
  } finally {
    clearTimeout(timeout);
  }
}

async function searchGoogleCseImages(query) {
  if (!GOOGLE_CSE_API_KEY || !GOOGLE_CSE_CX) {
    return [];
  }

  const endpoint = new URL("https://www.googleapis.com/customsearch/v1");
  endpoint.searchParams.set("key", GOOGLE_CSE_API_KEY);
  endpoint.searchParams.set("cx", GOOGLE_CSE_CX);
  endpoint.searchParams.set("q", query);
  endpoint.searchParams.set("searchType", "image");
  endpoint.searchParams.set("safe", "active");
  endpoint.searchParams.set("num", "5");
  endpoint.searchParams.set("imgSize", "large");

  const payload = await fetchJsonWithTimeout(endpoint.toString());
  const items = Array.isArray(payload?.items) ? payload.items : [];
  return items
    .map((item) => sanitizeHttpUrl(item?.link))
    .filter(Boolean)
    .map((url) => ({ url, provider: "google-cse", query }));
}

async function searchOpenverseImages(query) {
  const endpoint = new URL("https://api.openverse.org/v1/images/");
  endpoint.searchParams.set("q", query);
  endpoint.searchParams.set("page_size", "12");
  endpoint.searchParams.set("mature", "false");

  const payload = await fetchJsonWithTimeout(endpoint.toString());
  const results = Array.isArray(payload?.results) ? payload.results : [];
  const candidates = [];
  for (const result of results) {
    const urls = [result?.url, result?.thumbnail, result?.detail_url]
      .map((v) => sanitizeHttpUrl(v))
      .filter(Boolean);
    for (const url of urls) {
      candidates.push({ url, provider: "openverse", query });
      if (candidates.length >= 12) {
        return candidates;
      }
    }
  }
  return candidates;
}

function unsplashSourceUrl(query) {
  const normalized = normalizeQuery(query);
  if (!normalized) return "";
  return `https://source.unsplash.com/1600x900/?${encodeURIComponent(normalized)}`;
}

async function gatherImageCandidates(query) {
  const candidates = [];
  const seen = new Set();
  const push = (candidate) => {
    const url = sanitizeHttpUrl(candidate?.url);
    if (!url || seen.has(url)) return;
    seen.add(url);
    candidates.push({ ...candidate, url });
  };

  const google = await searchGoogleCseImages(query);
  for (const candidate of google) push(candidate);

  const openverse = await searchOpenverseImages(query);
  for (const candidate of openverse) push(candidate);

  const unsplash = unsplashSourceUrl(query);
  if (unsplash) {
    push({ url: unsplash, provider: "unsplash-source", query });
  }

  return candidates;
}

async function validateImageCandidate(url, options) {
  const normalizedUrl = sanitizeHttpUrl(url);
  if (!normalizedUrl) {
    return null;
  }

  const preferRaster = Boolean(options?.preferRaster);
  const attempts = [
    { method: "HEAD", headers: {} },
    { method: "GET", headers: { Range: "bytes=0-2048" } },
  ];

  for (const attempt of attempts) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), IMAGE_SEARCH_TIMEOUT_MS);
    try {
      const response = await fetch(normalizedUrl, {
        method: attempt.method,
        headers: attempt.headers,
        redirect: "follow",
        signal: controller.signal,
      });
      if (!response.ok) {
        continue;
      }

      const finalUrl = sanitizeHttpUrl(response.url || normalizedUrl);
      const contentType = String(response.headers.get("content-type") || "").toLowerCase();
      const isImage = contentType.includes("image/");
      if (!isImage) {
        continue;
      }
      if (preferRaster && isSvgTypeOrUrl(contentType, finalUrl)) {
        continue;
      }
      try {
        response.body?.cancel();
      } catch (_error) {
        // ignore
      }
      return {
        url: finalUrl,
        contentType,
      };
    } catch (_error) {
      // try next attempt
    } finally {
      clearTimeout(timeout);
    }
  }

  return null;
}

async function resolvePresentationImage({ selectedRecommendation, userPrompt, slideContext, content }) {
  const imagePayload = toImagePayload(content);
  const preferRaster = true;

  if (imagePayload.url) {
    const validatedDirect = await validateImageCandidate(imagePayload.url, { preferRaster });
    if (validatedDirect) {
      return {
        url: validatedDirect.url,
        source: "direct",
        query: imagePayload.query || imagePayload.alt || "",
      };
    }
  }

  const queries = await generateImageQueries({
    selectedRecommendation,
    userPrompt,
    slideContext,
    imagePayload,
  });
  debugLog("Image query candidates", queries);

  for (const query of queries) {
    const candidates = await gatherImageCandidates(query);
    for (const candidate of candidates.slice(0, 10)) {
      const validated = await validateImageCandidate(candidate.url, { preferRaster });
      if (validated) {
        return {
          url: validated.url,
          source: candidate.provider || "search",
          query,
        };
      }
    }
  }

  return null;
}

module.exports = {
  resolvePresentationImage,
};
