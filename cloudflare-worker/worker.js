/**
 * Cloudflare Worker — Gemini Vision Proxy
 * Receives a base64 image from index.html, calls Google Gemini,
 * and returns JSON with { individuals, sqor, named, horizon }.
 *
 * Required environment variable (set in the Cloudflare dashboard):
 *   GEMINI_API_KEY  →  your AIza... key from Google AI Studio
 */

const GEMINI_MODEL = "gemini-3.5-flash-lite";
const GEMINI_URL   = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

// Minimal prompt — fewer tokens, same accuracy
const PROMPT = `Extract these 6 numbers from the dashboard image. Convert $M/$K to full integers (e.g.$2.1M=2100000,$74K=74000). Use null if not found.
1.Attendees 2.SQO Creation 3.Named Accounts 4.Horizon Accounts 5.Event Cost 6.Wins`;

// Schema without descriptions — saves ~80 tokens per call
const RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    individuals: { type: "integer", nullable: true },
    sqor:        { type: "number",  nullable: true },
    named:       { type: "integer", nullable: true },
    horizon:     { type: "integer", nullable: true },
    cost:        { type: "number",  nullable: true },
    wins:        { type: "number",  nullable: true },
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
