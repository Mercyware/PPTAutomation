const OLLAMA_BASE_URL = process.env.OLLAMA_BASE_URL;
const OLLAMA_MODEL = process.env.OLLAMA_MODEL;

const AZURE_OPENAI_ENDPOINT = process.env.AZURE_OPENAI_ENDPOINT;
const AZURE_OPENAI_API_KEY = process.env.AZURE_OPENAI_API_KEY;
const AZURE_OPENAI_API_VERSION = process.env.AZURE_OPENAI_API_VERSION;
const AZURE_OPENAI_CHAT_DEPLOYMENT = process.env.AZURE_OPENAI_CHAT_DEPLOYMENT;
const AZURE_OPENAI_CHAT_MODEL = process.env.AZURE_OPENAI_CHAT_MODEL;
const LLM_TEMPERATURE = Number(process.env.LLM_TEMPERATURE);
const LLM_MAX_TOKENS = Number(process.env.LLM_MAX_TOKENS);

function isAzureConfigured() {
  return Boolean(
    AZURE_OPENAI_ENDPOINT &&
      AZURE_OPENAI_API_KEY &&
      AZURE_OPENAI_API_VERSION &&
      AZURE_OPENAI_CHAT_DEPLOYMENT
  );
}

function isOllamaConfigured() {
  return Boolean(OLLAMA_BASE_URL && OLLAMA_MODEL);
}

function assertConfigured() {
  if (isAzureConfigured()) {
    if (!Number.isFinite(LLM_TEMPERATURE)) {
      throw new Error("Invalid or missing LLM_TEMPERATURE in .env");
    }
    if (!Number.isFinite(LLM_MAX_TOKENS) || LLM_MAX_TOKENS <= 0) {
      throw new Error("Invalid or missing LLM_MAX_TOKENS in .env");
    }
    return;
  }

  if (isOllamaConfigured()) {
    return;
  }

  throw new Error(
    "LLM provider is not configured in .env. Configure either Azure OpenAI (AZURE_OPENAI_*) or Ollama (OLLAMA_BASE_URL + OLLAMA_MODEL)."
  );
}

function getAzureChatCompletionsUrl() {
  const baseUrl = AZURE_OPENAI_ENDPOINT.replace(/\/+$/, "");
  const deployment = encodeURIComponent(AZURE_OPENAI_CHAT_DEPLOYMENT);
  const apiVersion = encodeURIComponent(AZURE_OPENAI_API_VERSION);
  return `${baseUrl}/openai/deployments/${deployment}/chat/completions?api-version=${apiVersion}`;
}

function extractAzureMessageContent(payload) {
  const content = payload?.choices?.[0]?.message?.content;
  if (typeof content === "string") {
    return content;
  }

  if (Array.isArray(content)) {
    return content
      .map((part) => (typeof part?.text === "string" ? part.text : ""))
      .join("");
  }

  return "";
}

async function generateWithAzure({ systemPrompt, userPrompt, temperature }) {
  const requestBody = {
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
    max_completion_tokens: LLM_MAX_TOKENS,
    response_format: { type: "json_object" },
  };
  if (AZURE_OPENAI_CHAT_MODEL && AZURE_OPENAI_CHAT_MODEL.trim()) {
    requestBody.model = AZURE_OPENAI_CHAT_MODEL.trim();
  }

  const response = await fetch(getAzureChatCompletionsUrl(), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "api-key": AZURE_OPENAI_API_KEY,
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`Azure OpenAI request failed (${response.status}): ${errorBody}`);
  }

  const payload = await response.json();
  const content = extractAzureMessageContent(payload);
  if (!content) {
    throw new Error("Azure OpenAI returned an empty response.");
  }

  return content;
}

async function generateWithOllama({ systemPrompt, userPrompt, temperature }) {
  const response = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: OLLAMA_MODEL,
      prompt: `${systemPrompt}\n\nUser request:\n${userPrompt}`,
      stream: false,
      options: {
        temperature,
      },
      format: "json",
    }),
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(`Ollama request failed (${response.status}): ${errorBody}`);
  }

  const payload = await response.json();
  return payload.response;
}

async function generateStructuredJson({ systemPrompt, userPrompt, temperature = LLM_TEMPERATURE }) {
  assertConfigured();
  if (isAzureConfigured()) {
    return generateWithAzure({ systemPrompt, userPrompt, temperature });
  }

  return generateWithOllama({ systemPrompt, userPrompt, temperature });
}

const ACTIVE_PROVIDER = isAzureConfigured() ? "azure-openai" : isOllamaConfigured() ? "ollama" : "unconfigured";
const ACTIVE_MODEL = isAzureConfigured() ? AZURE_OPENAI_CHAT_MODEL : isOllamaConfigured() ? OLLAMA_MODEL : null;

module.exports = {
  generateStructuredJson,
  ACTIVE_PROVIDER,
  ACTIVE_MODEL,
};
