/**
 * Cloudflare Worker — Gemini Vision Proxy
 * Recebe uma imagem em base64 do index.html, chama o Google Gemini,
 * e devolve JSON com { individuals, sqor, named, horizon }.
 *
 * Variável de ambiente necessária (configurar no painel do Cloudflare):
 *   GEMINI_API_KEY  →  sua chave AIza... do Google AI Studio
 */

const GEMINI_MODEL = "gemini-3.6-flash";
const GEMINI_URL   = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

const PROMPT = `Look at this image. Find and return these four numbers:
1. The integer next to "Attendees" or "Total Attendees"
2. The dollar amount next to "SQO Creation" — convert to full number (e.g. $2.1M = 2100000, $0.7M = 700000, $1.5K = 1500). Return the full integer, no $ or suffixes.
3. The integer next to "Named Accounts"
4. The integer next to "Horizon Accounts"
If a value is not visible in the image, use null.`;

// Force Gemini to return a strict JSON schema — most reliable approach
const RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    individuals: { type: "integer", nullable: true, description: "Number next to Attendees or Total Attendees" },
    sqor:        { type: "number",  nullable: true, description: "Full dollar amount next to SQO Creation converted to integer (e.g. $2.1M = 2100000)" },
    named:       { type: "integer", nullable: true, description: "Number next to Named Accounts" },
    horizon:     { type: "integer", nullable: true, description: "Number next to Horizon Accounts" },
  },
  required: ["individuals", "sqor", "named", "horizon"],
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
      const err = await geminiResp.json().catch(() => ({}));
      const msg = err?.error?.message || `Gemini HTTP ${geminiResp.status}`;
      return jsonError(msg, 502);
    }

    const geminiData = await geminiResp.json();
    const rawText = geminiData?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";

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
