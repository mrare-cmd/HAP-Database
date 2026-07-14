/**
 * HAP DATABASE - update relay (Cloudflare Worker)
 *
 * Why this exists: GitHub requires a login for anything that writes, and a
 * token can't be put in the public page (GitHub auto-revokes any token it
 * finds published in a public repo). This tiny relay holds the token as a
 * secret. The page's green "Update to Current Month" button POSTs here,
 * and this triggers the GitHub Action that runs the unchanged HAP updater.
 * Result: ANYONE can press Update - no GitHub account needed.
 *
 * Setup: see WORKER_SETUP.txt. Requires two settings on the Worker:
 *   REPO          (plaintext variable)  e.g.  yourname/hap-database
 *   GITHUB_TOKEN  (secret)              fine-grained PAT, Actions: write
 *
 * A 3-minute cooldown blocks accidental double-clicks / spam. The workflow
 * itself also has a concurrency group, so runs can never overlap.
 */

let lastTrigger = 0; // per-isolate cooldown (best effort)

export default {
  async fetch(request, env) {
    const cors = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      "Content-Type": "application/json",
    };
    if (request.method === "OPTIONS")
      return new Response(null, { headers: cors });
    if (request.method !== "POST")
      return new Response(JSON.stringify({ ok: false, error: "POST only" }),
                          { status: 405, headers: cors });

    const now = Date.now();
    if (now - lastTrigger < 3 * 60 * 1000)
      return new Response(JSON.stringify({
        ok: true, note: "an update was already requested moments ago; it is running"
      }), { headers: cors });

    const r = await fetch(
      `https://api.github.com/repos/${env.REPO}/actions/workflows/update.yml/dispatches`,
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${env.GITHUB_TOKEN}`,
          "Accept": "application/vnd.github+json",
          "X-GitHub-Api-Version": "2022-11-28",
          "User-Agent": "hap-update-relay",
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ ref: "main" }),
      }
    );

    if (r.status === 204) {
      lastTrigger = now;
      return new Response(JSON.stringify({ ok: true }), { headers: cors });
    }
    const detail = await r.text().catch(() => "");
    return new Response(JSON.stringify({
      ok: false, error: `GitHub returned ${r.status}`, detail: detail.slice(0, 300)
    }), { status: 502, headers: cors });
  },
};
