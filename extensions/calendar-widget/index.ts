/**
 * Calendar widget for pi — shows your next meetings today, one line each,
 * rendered below the input editor (above the cwd/branch footer).
 *
 * Timing:
 *   - Microsoft Graph is polled every 5 minutes (meeting data rarely changes).
 *   - The countdown text is re-rendered locally every 30 seconds from the
 *     cached data — no network involved.
 *   - Hidden entirely when nothing is left today.
 *
 * Live data comes from Microsoft Graph (client credentials flow):
 *   O365_TENANT_ID, O365_CLIENT_ID, O365_CLIENT_SECRET, O365_UPN (optional,
 *   defaults to ska@vertica.dk), O365_TZ (optional, defaults to
 *   Europe/Copenhagen — the sandbox VM runs UTC, so the zone is configured,
 *   not detected). Without the credential variables the extension falls
 *   back to mock data so it can be tried out safely.
 *
 * Files:
 *   graph.ts  — pure Graph/parsing logic (no pi imports), unit-testable
 *   index.ts  — pi extension: widget rendering + timers
 */

import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import {
  buildWidgetLines,
  fetchTodayEvents,
  oauthConfigFromEnv,
  type GraphEvent,
} from "./graph.ts";

const FETCH_MS = 5 * 60_000; // poll Graph
const RENDER_MS = 30_000; // re-render countdowns locally
const MAX_EVENTS = 2;
// Must match graph.ts's default: the VM runs UTC, so the user's zone is
// configured, not detected. Keep in sync with oauthConfigFromEnv().
const DEFAULT_TZ = "Europe/Copenhagen";

/** Demo data mirroring a typical work day, anchored around "now" so the
 *  widget always has something to show regardless of when you launch pi.
 *  Times are rendered in the user's zone (tz), matching live data. */
function mockEvents(now: Date, tz: string): GraphEvent[] {
  // "minutes from now" expressed as wall-clock in tz: build the instant,
  // then take the naive local-time string the way Graph would return it
  // (with timeZone attached so parseGraphDateTime round-trips exactly).
  const fromNow = (minutes: number): { dateTime: string; timeZone: string } => {
    const d = new Date(now.getTime() + minutes * 60_000);
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: tz,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    }).formatToParts(d);
    const p = (t: string) => parts.find((x) => x.type === t)?.value ?? "00";
    return {
      dateTime: `${p("year")}-${p("month")}-${p("day")}T${p("hour")}:${p("minute")}:${p("second")}.000`,
      timeZone: tz,
    };
  };
  const ev = (
    subject: string,
    startMinFromNow: number,
    minutes: number,
    location: string,
    online = false,
  ): GraphEvent => ({
    subject,
    start: fromNow(startMinFromNow),
    end: fromNow(startMinFromNow + minutes),
    location: { displayName: location },
    isOnlineMeeting: online,
    onlineMeetingProvider: online ? "teamsForBusiness" : null,
  });

  return [
    ev("Architecture review", -10, 90, "Lille Sal / 2.1"), // in progress
    ev("Lunch & learn: Rust", 40, 45, "Teams", true),
    ev("1:1 with manager", 110, 30, "Teams", true),
    ev("Customer workshop", 170, 120, "Store Sal / 2.2"),
  ];
}

export default function (pi: ExtensionAPI): void {
  let fetchTimer: ReturnType<typeof setInterval> | undefined;
  let renderTimer: ReturnType<typeof setInterval> | undefined;
  let usingMock = false;
  let tz = DEFAULT_TZ;
  // Cache of the last known event set; countdowns re-render from this.
  let cachedEvents: GraphEvent[] | undefined;

  function render(ctx: ExtensionContext) {
    const now = new Date();
    const events = usingMock || !cachedEvents ? mockEvents(now, tz) : cachedEvents;
    // Urgent countdowns (next meeting ≤ 5 min away) render in the theme's error color.
    const colorize = (text: string, urgent: boolean) => (urgent ? ctx.ui.theme.fg("error", text) : text);
    const lines = buildWidgetLines(events, now, MAX_EVENTS, colorize, tz);
    if (lines.length === 0) {
      ctx.ui.setWidget("calendar", undefined);
      return;
    }
    if (usingMock) lines.push("~ mock data (set O365_* env vars) ~");
    ctx.ui.setWidget("calendar", lines, { placement: "belowEditor" });
  }

  function fetchNow(ctx: ExtensionContext) {
    if (usingMock) {
      render(ctx); // mock events are relative to now — refresh them too
      return;
    }
    const cfg = oauthConfigFromEnv();
    if (!cfg) {
      usingMock = true;
      render(ctx);
      return;
    }
    tz = cfg.tz;
    fetchTodayEvents(cfg, new Date())
      .then((events) => {
        cachedEvents = events;
        usingMock = false;
        render(ctx);
      })
      .catch((err: unknown) => {
        usingMock = true;
        render(ctx);
        ctx.ui.notify(
          `Calendar: ${err instanceof Error ? err.message : String(err)} — showing mock data`,
          "warning",
        );
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
