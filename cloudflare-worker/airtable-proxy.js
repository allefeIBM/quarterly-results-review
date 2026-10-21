/**
 * Cloudflare Worker — Airtable Proxy
 *
 * Proxies Airtable API requests from the browser so the token never
 * needs to be exposed to end-users.
 *
 * Required environment variable (set in the Cloudflare dashboard as a Secret):
 *   AIRTABLE_TOKEN  →  your patXXX... personal access token from airtable.com/create/tokens
 *
 * Supported endpoints (called by index.html):
 *   POST /search   { q: "<search string>" }
 *     → searches by event name OR brief code, returns [{event_name, brief}]
 *
 *   POST /brief    { brief: "<BRIEF_CODE>" }
 *     → returns full detail for one event: { event_name, event_date, expected_individuals, expected_sqor_bucket }
 */

// ── Airtable coordinates (public — no secrets here) ────────────────────────
const AT_BASE_ID    = "appcSYlaIYXL0UHy9";
const AT_TABLE_ID   = "tblFRyAdPMPYwU8Zr";
const AT_BRIEF_FIELD    = "🔷EMB #";
const AT_NAME_FIELD     = "🔷Event Name";
const AT_EXPECTED_FIELD = "🔹Expected Responses";
const AT_DATE_FIELD     = "🔷Start Date";

const AT_BASE_URL = `https://api.airtable.com/v0/${AT_BASE_ID}/${AT_TABLE_ID}`;

// ── Entry point ────────────────────────────────────────────────────────────
export default {
  async fetch(request, env) {

    // CORS preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }

    if (request.method !== "POST") {
      return jsonError("Method not allowed", 405);
    }

    if (!env.AIRTABLE_TOKEN) {
      return jsonError("AIRTABLE_TOKEN not configured in Worker environment", 500);
    }

    const url = new URL(request.url);
    let body;
    try {
      body = await request.json();
    } catch {
      return jsonError("Invalid JSON body", 400);
    }

    const atHeaders = {
      "Authorization": `Bearer ${env.AIRTABLE_TOKEN}`,
      "Content-Type": "application/json",
    };

    // ── POST /search ────────────────────────────────────────────────────────
    if (url.pathname.endsWith("/search")) {
      const q = (body.q || "").trim();
      if (!q) return jsonError("Missing field: q", 400);

      const safeQ = q.replace(/\\/g, "\\\\").replace(/"/g, '\\"').toLowerCase();
      const filter = encodeURIComponent(
        `OR(FIND("${safeQ}",LOWER({${AT_NAME_FIELD}}&"")),FIND("${safeQ}",LOWER({${AT_BRIEF_FIELD}}&"")))`
      );
      const atUrl = `${AT_BASE_URL}?filterByFormula=${filter}&maxRecords=30` +
        `&fields[]=${encodeURIComponent(AT_NAME_FIELD)}&fields[]=${encodeURIComponent(AT_BRIEF_FIELD)}`;

      const atResp = await fetch(atUrl, { headers: atHeaders });
      if (!atResp.ok) return jsonError(`Airtable error ${atResp.status}`, 502);

      const data = await atResp.json();
      const results = (data.records || [])
        .map(r => ({
          event_name: r.fields[AT_NAME_FIELD] || "",
          brief: (r.fields[AT_BRIEF_FIELD] || "").trim(),
        }))
        .filter(r => r.brief);

      return new Response(JSON.stringify({ results }), {
        status: 200,
        headers: { "Content-Type": "application/json", ...corsHeaders() },
      });
    }

    // ── POST /brief ─────────────────────────────────────────────────────────
    if (url.pathname.endsWith("/brief")) {
      const brief = (body.brief || "").trim();
      if (!brief) return jsonError("Missing field: brief", 400);

      const filter = encodeURIComponent(`{${AT_BRIEF_FIELD}}="${brief}"`);
      const atUrl = `${AT_BASE_URL}?filterByFormula=${filter}&maxRecords=5`;

      const atResp = await fetch(atUrl, { headers: atHeaders });
      if (!atResp.ok) return jsonError(`Airtable error ${atResp.status}`, 502);

      const data = await atResp.json();
      const record = data.records && data.records[0];
      if (!record) {
        return new Response(JSON.stringify({ result: null }), {
          status: 200,
          headers: { "Content-Type": "application/json", ...corsHeaders() },
        });
      }

      const fields = record.fields || {};
      const eventDate = parseDate(fields[AT_DATE_FIELD]);
      const result = {
        event_name:           fields[AT_NAME_FIELD] || "",
        event_date:           eventDate || "",
        expected_individuals: fields[AT_EXPECTED_FIELD] ?? null,
        expected_sqor_bucket: quarterBucket(eventDate),
      };

      return new Response(JSON.stringify({ result }), {
        status: 200,
        headers: { "Content-Type": "application/json", ...corsHeaders() },
      });
    }

    return jsonError("Not found", 404);
  },
};

// ── Helpers ────────────────────────────────────────────────────────────────

function parseDate(value) {
  if (!value) return null;
  if (Array.isArray(value)) {
    for (const v of value) { const d = parseDate(v); if (d) return d; }
    return null;
  }
  const text = String(value).trim().replace("Z", "+00:00");
  const attempts = [
    () => new Date(text),
    () => { const [m, d, y] = text.split("/"); return new Date(y, m - 1, d); },
  ];
  for (const fn of attempts) {
    try { const d = fn(); if (!isNaN(d)) return d.toISOString().split("T")[0]; } catch (_) {}
  }
  return null;
}

function quarterBucket(isoDate) {
  if (!isoDate) return null;
  const eventDate = new Date(isoDate);
  if (isNaN(eventDate)) return null;
  const now = new Date();
  const months = (now.getFullYear() - eventDate.getFullYear()) * 12 +
                 (now.getMonth() - eventDate.getMonth());
  const q = Math.floor(months / 3);
  if (q <= 0) return "CQ";
  if (q === 1) return "NQ";
  if (q === 2) return "NQ+1";
  return "NQ+2";
}

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
