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

const PROMPT = `Look at this image and find these four numbers. Return ONLY a single JSON object, nothing else — no text before, no text after, no markdown, no explanation.

JSON keys to return:
"individuals" = the integer next to the label "Attendees" or "Total Attendees" (null if not visible)
"sqor" = the number next to the label "SQO Creation" with $ removed and no commas (null if not visible)
"named" = the integer next to the label "Named Accounts" (null if not visible)
"horizon" = the integer next to the label "Horizon Accounts" (null if not visible)

Your entire response must be exactly this format and nothing else:
{"individuals":245,"sqor":180000,"named":120,"horizon":55}`;

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
        generationConfig: { maxOutputTokens: 200, temperature: 0 },
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
