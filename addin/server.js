const path = require("path");
const https = require("https");
const express = require("express");
const devCerts = require("office-addin-dev-certs");

const PORT = Number(process.env.ADDIN_PORT || 3100);
const BACKEND_BASE_URL = process.env.BACKEND_BASE_URL || "http://localhost:4000";

async function start() {
  const app = express();
  const webRoot = path.join(__dirname, "web");
  const srcRoot = path.join(__dirname, "src");

  app.use((_req, res, next) => {
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    next();
  });

  app.use(express.json({ limit: "3mb" }));
  app.use(express.static(webRoot));
  app.use("/src", express.static(srcRoot));

  app.post("/api/recommendations", async (req, res) => {
    try {
      const upstream = await fetch(`${BACKEND_BASE_URL}/api/recommendations`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(req.body || {}),
      });
      const body = await upstream.text();
      res.status(upstream.status).type("application/json").send(body);
    } catch (error) {
      res.status(502).json({
        error: "Backend proxy failed for recommendations",
        details: error && error.message ? error.message : String(error),
      });
    }
  });

  app.post("/api/plans", async (req, res) => {
    try {
      const upstream = await fetch(`${BACKEND_BASE_URL}/api/plans`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(req.body || {}),
      });
      const body = await upstream.text();
      res.status(upstream.status).type("application/json").send(body);
    } catch (error) {
      res.status(502).json({
        error: "Backend proxy failed for plans",
        details: error && error.message ? error.message : String(error),
      });
    }
  });

  app.get("/health", (_req, res) => {
    res.json({ ok: true, service: "ppt-automation-addin-web", port: PORT });
  });

  const httpsOptions = await devCerts.getHttpsServerOptions();
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`addin web host running at https://localhost:${PORT}`);
  });
}

start().catch((error) => {
  console.error("Failed to start add-in web host:", error);
  process.exit(1);
});
