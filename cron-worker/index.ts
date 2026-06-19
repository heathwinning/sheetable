/**
 * Cron Worker for Sheetable snapshot scheduling.
 * Deployed separately from the Pages app. Uses Cloudflare Workers Cron Triggers
 * to call the Pages /api/cron/snapshots endpoint on a schedule.
 *
 * Deploy: npx wrangler deploy --config cron-worker/wrangler.toml
 */

interface Env {
  PAGES_URL: string;
  CRON_SECRET: string;
}

export default {
  async scheduled(_controller: ScheduledController, env: Env, _ctx: ExecutionContext): Promise<void> {
    const url = `${env.PAGES_URL}/api/cron/snapshots`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${env.CRON_SECRET}` },
    });

    if (!resp.ok) {
      const body = await resp.text().catch(() => '');
      throw new Error(`Cron snapshot failed: HTTP ${resp.status} ${body.slice(0, 200)}`);
    }

    const result = await resp.json() as { created: number };
    console.log(`Snapshot cron complete: ${result.created} snapshots created`);
  },
};
