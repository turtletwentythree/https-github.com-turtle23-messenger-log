/**
 * messenger-sync-trigger
 *
 * Small public-facing proxy Worker that lets the (fully public, static)
 * Messenger log site trigger a fresh SharePoint -> booking_data.json sync
 * on demand, without ever exposing a GitHub token in client-side code.
 *
 * Flow:
 *   1. Browser POSTs here (no body needed) when the person clicks
 *      "ซิงค์ข้อมูลล่าสุด".
 *   2. This Worker checks a simple KV-backed rate limit (one trigger per
 *      MIN_INTERVAL_MS across all visitors) to avoid hammering the
 *      GitHub Actions workflow / SharePoint.
 *   3. If allowed, it calls GitHub's workflow_dispatch API for
 *      sync.yml using a GITHUB_TOKEN held only as a Worker secret.
 *   4. The GitHub Action itself pulls fresh data from SharePoint and
 *      commits an updated booking_data.json with a new generated_at
 *      timestamp; the page then polls for that file to change.
 */

const ALLOWED_ORIGIN = "https://turtletwentythree.github.io";
const REPO = "turtletwentythree/https-github.com-turtle23-messenger-log";
const WORKFLOW_FILE = "sync.yml";
const REF = "main";
const MIN_INTERVAL_MS = 60 * 1000; // one trigger per minute, across all visitors
const KV_KEY = "last_trigger_at";

function corsHeaders(origin) {
  return {
    "Access-Control-Allow-Origin": origin === ALLOWED_ORIGIN ? origin : ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400",
  };
}

function json(data, status, extraHeaders) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...extraHeaders,
    },
  });
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(origin);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: cors });
    }

    if (request.method !== "POST") {
      return json({ ok: false, error: "method_not_allowed" }, 405, cors);
    }

    if (!env.SYNC_STATE || !env.GH_PAT) {
      return json({ ok: false, error: "worker_not_configured" }, 500, cors);
    }

    try {
      const now = Date.now();
      const lastStr = await env.SYNC_STATE.get(KV_KEY);
      const last = lastStr ? parseInt(lastStr, 10) : 0;
      const elapsedMs = now - (Number.isFinite(last) ? last : 0);

      if (elapsedMs < MIN_INTERVAL_MS) {
        const retryAfterSec = Math.max(1, Math.ceil((MIN_INTERVAL_MS - elapsedMs) / 1000));
        return json(
          { ok: false, rateLimited: true, retryAfterSec },
          429,
          { ...cors, "Retry-After": String(retryAfterSec) }
        );
      }

      // Reserve the slot before calling out, so concurrent clicks don't
      // both slip through while the GitHub call is in flight.
      await env.SYNC_STATE.put(KV_KEY, String(now));

      const ghResp = await fetch(
        `https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${env.GH_PAT}`,
            Accept: "application/vnd.github+json",
            "User-Agent": "messenger-sync-trigger-worker",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ ref: REF }),
        }
      );

      if (ghResp.status === 204) {
        return json({ ok: true, triggered: true, at: now }, 200, cors);
      }

      const detail = await ghResp.text();
      return json(
        {
          ok: false,
          error: "github_dispatch_failed",
          status: ghResp.status,
          detail: detail.slice(0, 400),
        },
        502,
        cors
      );
    } catch (err) {
      return json(
        { ok: false, error: "exception", message: String((err && err.message) || err) },
        500,
        cors
      );
    }
  },
};
