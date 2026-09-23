/**
 * Calendar widget for pi — shows your next meetings today, one line each,
 * rendered below the input editor (above the cwd/branch footer).
 *
 * Timing:
 *   - Microsoft Graph is polled every 5 minutes (meeting data rarely changes).
 *   - The countdown text is re-rendered locally every 30 seconds from the
 *     cached data — no network involved.
 *   - Shows "no meetings left" when nothing is left today.
 *
 * Live data comes from Microsoft Graph (client credentials flow):
 *   O365_TENANT_ID, O365_CLIENT_ID, O365_CLIENT_SECRET, O365_UPN (optional,
 *   defaults to ska@vertica.dk), O365_TZ (optional, defaults to
 *   Europe/Copenhagen — the sandbox VM runs UTC, so the zone is configured,
 *   not detected). Until the first successful fetch the widget shows "~"
 *   (no connection); a later failed fetch never invalidates data already
 *   fetched — the last good events keep rendering until Graph is reachable
 *   again.
 *
 * Files:
 *   graph.ts  — pure Graph/parsing logic (no pi imports), unit-testable
 *   index.ts  — pi extension: widget rendering + timers
 */

import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import { buildWidgetLines, fetchTodayEvents, oauthConfigFromEnv, type GraphEvent } from "./graph.ts";

const FETCH_MS = 5 * 60_000; // poll Graph
const RENDER_MS = 30_000; // re-render countdowns locally
const MAX_EVENTS = 2;
// Must match graph.ts's default: the VM runs UTC, so the user's zone is
// configured, not detected. Keep in sync with oauthConfigFromEnv().
const DEFAULT_TZ = "Europe/Copenhagen";

export default function (pi: ExtensionAPI): void {
  let fetchTimer: ReturnType<typeof setInterval> | undefined;
  let renderTimer: ReturnType<typeof setInterval> | undefined;
  let tz = DEFAULT_TZ;
  // Last successful fetch's events; undefined until the first fetch succeeds.
  // A failed fetch never clears this — the widget keeps rendering it.
  let cachedEvents: GraphEvent[] | undefined;

  function render(ctx: ExtensionContext) {
    if (!cachedEvents) {
      // Never fetched successfully yet (or no O365_* config at all): "~"
      // means "no connection", NOT "no meetings".
      ctx.ui.setWidget("calendar", ["~"], { placement: "belowEditor" });
      return;
    }
    const now = new Date();
    // Urgent countdowns (next meeting ≤ 5 min away) render in the theme's error color.
    const colorize = (text: string, urgent: boolean) => (urgent ? ctx.ui.theme.fg("error", text) : text);
    const lines = buildWidgetLines(cachedEvents, now, MAX_EVENTS, colorize, tz);
    if (lines.length > 0) {
      ctx.ui.setWidget("calendar", lines, { placement: "belowEditor" });
      return;
    }
    ctx.ui.setWidget("calendar", ["no meetings left"], { placement: "belowEditor" });
  }

  function fetchNow(ctx: ExtensionContext) {
    const cfg = oauthConfigFromEnv();
    if (!cfg) return; // no credentials — keep showing "~"
    tz = cfg.tz;
    fetchTodayEvents(cfg, new Date())
      .then((events) => {
        cachedEvents = events;
        render(ctx);
      })
      .catch((err: unknown) => {
        // Keep cachedEvents as-is: never invalidate good data on a failed
        // request — only the initial fetch can land on "~".
        render(ctx);
        ctx.ui.notify(`Calendar: ${err instanceof Error ? err.message : String(err)}`, "warning");
      });
  }

  pi.on("session_start", (_event, ctx) => {
    if (!ctx.hasUI) return;

    // Stop any previous timers (e.g. after /reload → session_start again).
    if (fetchTimer) clearInterval(fetchTimer);
    if (renderTimer) clearInterval(renderTimer);

    fetchNow(ctx);
    renderTimer = setInterval(() => render(ctx), RENDER_MS);
    fetchTimer = setInterval(() => fetchNow(ctx), FETCH_MS);
  });

  pi.on("session_shutdown", () => {
    if (fetchTimer) {
      clearInterval(fetchTimer);
      fetchTimer = undefined;
    }
    if (renderTimer) {
      clearInterval(renderTimer);
      renderTimer = undefined;
    }
  });
}
