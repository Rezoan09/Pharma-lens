// api/insights.js — PharmaLens API proxy with usage limits

const MODEL = "claude-sonnet-4-6";   // locked: the browser cannot choose a model
const MAX_TOKENS = 1000;             // locked: caps the cost of each call
const PER_IP_LIMIT = 5;              // analyses per IP per day
const GLOBAL_DAILY_LIMIT = 20;       // analyses per day across all users
const WINDOW_MS = 24 * 60 * 60 * 1000;

// In-memory counters (best effort: they reset when Vercel restarts the function)
const hits = new Map();
let day = { date: "", count: 0 };

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY || process.env.VITE_ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({ error: "Server is not configured" });
  }

  // Only accept PharmaLens-style requests, not free-form chat
  const body = req.body || {};
  const userMsg = body.messages?.[0]?.content;
  const system = typeof body.system === "string" ? body.system : "";
  if (
    typeof userMsg !== "string" ||
    !userMsg.startsWith("Analyze this") ||
    userMsg.length > 8000 ||
    system.length > 2000
  ) {
    return res.status(400).json({ error: "Invalid request" });
  }

  // Per-IP limit
  const ip =
    (req.headers["x-forwarded-for"] || "").split(",")[0].trim() ||
    req.socket?.remoteAddress ||
    "unknown";
  const now = Date.now();
  const recent = (hits.get(ip) || []).filter((t) => now - t < WINDOW_MS);
  if (recent.length >= PER_IP_LIMIT) {
    return res.status(429).json({ error: "Daily demo limit reached for your connection (5 analyses per day). Please try again tomorrow or request a full demo." });
  }

  // Global daily limit
  const today = new Date().toISOString().slice(0, 10);
  if (day.date !== today) day = { date: today, count: 0 };
  if (day.count >= GLOBAL_DAILY_LIMIT) {
    return res.status(429).json({ error: "Today\u2019s free demo capacity is full. Please try again tomorrow or request a full demo." });
  }

  recent.push(now);
  hits.set(ip, recent);
  day.count += 1;

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: MAX_TOKENS,
        system,
        messages: [{ role: "user", content: userMsg }],
      }),
    });
    const data = await response.json();
    return res.status(response.status).json(data);
  } catch (err) {
    return res.status(502).json({ error: "Upstream request failed" });
  }
}
