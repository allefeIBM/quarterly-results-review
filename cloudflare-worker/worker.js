/**
 * Cloudflare Worker — Gemini Vision Proxy
 * Receives a base64 image from index.html, calls Google Gemini,
 * and returns JSON with { individuals, sqor, named, horizon }.
 *
 * Required environment variable (set in the Cloudflare dashboard):
 *   GEMINI_API_KEY  →  your AIza... key from Google AI Studio
 */

const GEMINI_MODEL = "gemini-2.0-flash";
const GEMINI_URL   = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

const PROMPT = `Look at this image. Find and return these numbers:
1. The integer next to "Attendees" or "Total Attendees"
2. The dollar amount next to "SQO Creation" — convert to full number (e.g. $2.1M = 2100000, $0.7M = 700000, $1.5K = 1500). Return the full integer, no $ or suffixes.
3. The integer next to "Named Accounts"
4. The integer next to "Horizon Accounts"
5. The dollar amount next to "Event Cost" or "Cost" — convert to full integer (e.g. $74,245 = 74245, $74.2K = 74200, $0.1M = 100000). Return the full integer, no $ or suffixes.
6. The dollar amount next to "Wins" or "Win" or "Won" — convert to full integer (e.g. $0.4M = 400000, $400K = 400000, $50,000 = 50000). Return the full integer, no $ or suffixes.
If a value is not visible in the image, use null.`;

// Force Gemini to return a strict JSON schema — most reliable approach
const RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    individuals: { type: "integer", nullable: true, description: "Number next to Attendees or Total Attendees" },
    sqor:        { type: "number",  nullable: true, description: "Full dollar amount next to SQO Creation converted to integer (e.g. $2.1M = 2100000)" },
    named:       { type: "integer", nullable: true, description: "Number next to Named Accounts" },
    horizon:     { type: "integer", nullable: true, description: "Number next to Horizon Accounts" },
    cost:        { type: "number",  nullable: true, description: "Full dollar amount next to Event Cost converted to integer (e.g. $74,245 = 74245)" },
    wins:        { type: "number",  nullable: true, description: "Full dollar amount next to Wins converted to integer (e.g. $0.4M = 400000)" },
  },
  required: ["individuals", "sqor", "named", "horizon", "cost", "wins"],
};

export default {
  async fetch(request, env) {

    // ── CORS preflight ──────────────────────────────────────────────────────
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders(),
      });
    }

    if (request.method !== "POST") {
      return new Response("Method not allowed", { status: 405, headers: corsHeaders() });
    }

    // ── Parse incoming body ─────────────────────────────────────────────────
    let body;
    try {
      body = await request.json();
    } catch {
      return jsonError("Invalid JSON body", 400);
    }

    const { image_base64, mime_type } = body;
    if (!image_base64 || !mime_type) {
      return jsonError("Missing image_base64 or mime_type", 400);
    }

    if (!env.GEMINI_API_KEY) {
      return jsonError("GEMINI_API_KEY not configured in Worker environment", 500);
    }

    // ── Call Gemini ─────────────────────────────────────────────────────────
    const geminiResp = await fetch(`${GEMINI_URL}?key=${env.GEMINI_API_KEY}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{
          parts: [
            { text: PROMPT },
            { inline_data: { mime_type, data: image_base64 } },
          ],
        }],
        generationConfig: {
          temperature: 0,
          responseMimeType: "application/json",
          responseSchema: RESPONSE_SCHEMA,
        },
      }),
    });

    if (!geminiResp.ok) {
      const errBody = await geminiResp.json().catch(() => ({}));
      const msg = errBody?.error?.message || `Gemini HTTP ${geminiResp.status}`;
      const detail = errBody?.error?.status || "";
      return jsonError(`${msg}${detail ? " (" + detail + ")" : ""}`, 502);
    }

    const geminiData = await geminiResp.json();
    const rawText = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";

    if (!rawText) {
      return jsonError("Gemini returned an empty response — image may be unsupported or blocked", 502);
    }

    // Return raw text — index.html will parse the JSON
    return new Response(JSON.stringify({ result: rawText }), {
      status: 200,
      headers: { "Content-Type": "application/json", ...corsHeaders() },
    });
  },
};

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function jsonError(message, status) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() },
  });
}
