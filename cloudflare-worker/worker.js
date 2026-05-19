/**
 * Cardiotrack sync trigger — Cloudflare Worker.
 *
 * Receives a POST from the public dashboard and triggers the GitHub Actions
 * workflow that pulls fresh Drive + Gmail data and redeploys the site.
 *
 * Secrets (set via `wrangler secret put` — see SETUP.md):
 *   GH_TOKEN    — fine-grained PAT, repo scope: Actions write, Contents read,
 *                 Workflows write. ONLY scoped to the cardiotrack-dashboard repo.
 *   GH_REPO     — "<user>/<repo>", e.g. "islamxnajmul/cardiotrack-dashboard"
 *   WORKFLOW    — workflow filename, default "sync.yml"
 *
 * Endpoint:
 *   POST /        → triggers the workflow, returns {ok, run_url, message}
 *   GET  /status  → returns latest run state {status, conclusion, html_url}
 *   GET  /        → health-check
 */

const CORS = {
  "Access-Control-Allow-Origin":  "*",   // tighten to your Pages URL in production
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...CORS },
  });
}

async function triggerWorkflow(env) {
  const workflow = env.WORKFLOW || "sync.yml";
  const url = `https://api.github.com/repos/${env.GH_REPO}/actions/workflows/${workflow}/dispatches`;

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${env.GH_TOKEN}`,
      "Accept":        "application/vnd.github+json",
      "User-Agent":    "cardiotrack-sync-trigger/1.0",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    body: JSON.stringify({ ref: env.GH_BRANCH || "main" }),
  });

  // GitHub returns 204 No Content on success
  if (resp.status === 204) {
    return { ok: true };
  }
  const text = await resp.text();
  return { ok: false, status: resp.status, error: text.slice(0, 400) };
}

async function getLatestRun(env) {
  const workflow = env.WORKFLOW || "sync.yml";
  const url = `https://api.github.com/repos/${env.GH_REPO}/actions/workflows/${workflow}/runs?per_page=1`;
  const resp = await fetch(url, {
    headers: {
      "Authorization": `Bearer ${env.GH_TOKEN}`,
      "Accept":        "application/vnd.github+json",
      "User-Agent":    "cardiotrack-sync-trigger/1.0",
    },
  });
  if (!resp.ok) {
    return { ok: false, status: resp.status, error: (await resp.text()).slice(0, 400) };
  }
  const body = await resp.json();
  const run  = body.workflow_runs?.[0];
  if (!run) return { ok: true, run: null };
  return {
    ok: true,
    run: {
      id:           run.id,
      status:       run.status,         // queued | in_progress | completed
      conclusion:   run.conclusion,     // success | failure | null
      html_url:     run.html_url,
      created_at:   run.created_at,
      updated_at:   run.updated_at,
      run_started_at: run.run_started_at,
      event:        run.event,
    },
  };
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: CORS });
    }

    const url = new URL(request.url);

    // Health-check + simple landing page
    if (request.method === "GET" && url.pathname === "/") {
      return json({
        service: "cardiotrack-sync-trigger",
        repo: env.GH_REPO || "(not configured)",
        endpoints: { trigger: "POST /", status: "GET /status" },
      });
    }

    // Get latest run state (used by the dashboard's polling loop)
    if (request.method === "GET" && url.pathname === "/status") {
      if (!env.GH_TOKEN || !env.GH_REPO) {
        return json({ ok: false, error: "Worker missing GH_TOKEN or GH_REPO secret" }, 500);
      }
      return json(await getLatestRun(env));
    }

    // Trigger a fresh sync
    if (request.method === "POST" && url.pathname === "/") {
      if (!env.GH_TOKEN || !env.GH_REPO) {
        return json({ ok: false, error: "Worker missing GH_TOKEN or GH_REPO secret" }, 500);
      }
      const result = await triggerWorkflow(env);
      if (!result.ok) {
        return json({ ok: false, ...result }, 502);
      }
      return json({
        ok: true,
        message: "Workflow triggered. Poll /status to watch progress.",
        actions_url: `https://github.com/${env.GH_REPO}/actions/workflows/${env.WORKFLOW || "sync.yml"}`,
      });
    }

    return json({ ok: false, error: "Not found" }, 404);
  },
};
