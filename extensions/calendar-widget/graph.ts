/**
 * Pure (UI-free) Microsoft Graph calendar logic for the calendar widget.
 *
 * This module intentionally has zero imports so it can be unit-tested
 * directly with `node --experimental-strip-types` without loading pi.
 */

export interface GraphEvent {
  subject?: string | null;
  start?: { dateTime?: string | null; timeZone?: string | null } | null;
  end?: { dateTime?: string | null; timeZone?: string | null } | null;
  location?: { displayName?: string | null } | null;
  isOnlineMeeting?: boolean | null;
  onlineMeetingProvider?: string | null;
  bodyPreview?: string | null;
}

export interface WidgetEvent {
  subject: string;
  location: string; // room name or "Teams"
  startsAt: Date;
  endsAt: Date;
  durationMs: number;
}

const TEAMS_PROVIDERS = new Set(["teamsForBusiness", "skypeForBusiness", "unknown"]);

// ---------------------------------------------------------------------------
// Timezone helpers
//
// The sandbox VM runs UTC, but meetings live in the user's local zone
// (Europe/Copenhagen). Graph's calendarView returns naive wall-clock strings
// ("2026-09-23T14:00:00.000") qualified by start/end.timeZone (a Windows
// name like "W. Europe Standard Time" or "UTC"). We therefore:
//   1. ask Graph for UTC ("Prefer: outlook.timezone=UTC" + buildDayWindow),
//   2. convert the wall clock from the event's zone to a real instant, and
//   3. filter/render in the user's zone via Intl (DST-safe, no tz database
//      dependency — Node has full ICU).
// ---------------------------------------------------------------------------

/** Map an IANA zone to its current UTC offset ("+02:00"), or null if unknown. */
export function timeZoneOffsetMinutes(tz: string, at: Date): number | null {
  try {
    const dtf = new Intl.DateTimeFormat("en-US", {
      timeZone: tz,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hourCycle: "h23",
    });
    const parts = dtf.formatToParts(at);
    const get = (type: string) => Number(parts.find((p) => p.type === type)?.value);
    const asUTC = Date.UTC(
      get("year"),
      get("month") - 1,
      get("day"),
      get("hour") % 24,
      get("minute"),
      get("second"),
    );
    return Math.round((asUTC - at.getTime()) / 60_000);
  } catch {
    return null;
  }
}

/**
 * Interpret `dt` ("yyyy-MM-ddTHH:mm:ss[.SSS][Z|±hh:mm]") as wall-clock time in
 * `tz` and return the corresponding instant. Naive strings from Graph have no
 * offset, so `new Date()` would wrongly assume the VM's zone (UTC in sandboxes).
 */
export function parseGraphDateTime(dt: string | null | undefined, tz = "UTC"): Date | null {
  if (!dt) return null;
  if (/[zZ]|[+-]\d{2}:?\d{2}$/.test(dt.trim())) {
    // Carries an explicit offset/Z — the ISO parser handles it correctly
    // regardless of the host zone.
    const d = new Date(dt);
    return Number.isNaN(d.getTime()) ? null : d;
  }
  // Naive wall clock: anchor it to UTC (host-zone independent reference —
  // `new Date(naive)` would parse it in the HOST's zone, which breaks on
  // non-UTC hosts), then subtract the target zone's offset:
  //   wall 14:00 in UTC+1  =>  instant 13:00Z.
  const m = dt.trim().match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?(?:\.(\d{1,3}))?$/);
  if (!m) return null;
  const [, y, mo, da, h, mi, s = "0", ms = "0"] = m;
  const asUTC = Date.UTC(Number(y), Number(mo) - 1, Number(da), Number(h), Number(mi), Number(s), Number(ms.padEnd(3, "0")));
  const offsetMin = timeZoneOffsetMinutes(tz, new Date(asUTC));
  if (offsetMin === null) return null; // unknown tz — caller falls back
  return new Date(asUTC - offsetMin * 60_000);
}

/** "2026-09-23" — the calendar day of `d` as seen in `tz` (DST-safe). */
export function wallDateInZone(d: Date, tz: string): string {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: tz,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(d);
}

/** Extract today's upcoming events (not already ended), sorted by start time.
 *  "Today" is evaluated in `tz` (the user's zone), not the VM's. */
export function pickNextEvents(
  events: GraphEvent[],
  now: Date,
  count: number,
  tz = "UTC",
): WidgetEvent[] {
  const today = wallDateInZone(now, tz);
  const upcoming: WidgetEvent[] = [];

  for (const e of events) {
    // Prefer the event's own timeZone; fall back to the user's zone.
    const start = parseGraphDateTime(e.start?.dateTime, e.start?.timeZone || tz) ??
      parseGraphDateTime(e.start?.dateTime, tz);
    const end = parseGraphDateTime(e.end?.dateTime, e.end?.timeZone || tz) ??
      parseGraphDateTime(e.end?.dateTime, tz);
    if (!start || !end) continue;
    if (wallDateInZone(start, tz) !== today) continue;
    if (end.getTime() <= now.getTime()) continue;
    upcoming.push({
      subject: e.subject?.trim() || "(no title)",
      location: resolveLocation(e),
      startsAt: start,
      endsAt: end,
      durationMs: end.getTime() - start.getTime(),
    });
  }

  upcoming.sort((a, b) => a.startsAt.getTime() - b.startsAt.getTime());
  return upcoming.slice(0, count);
}

function resolveLocation(e: GraphEvent): string {
  // Online-only meetings (Teams etc.) win over the room name.
  if (e.isOnlineMeeting && (!e.onlineMeetingProvider || TEAMS_PROVIDERS.has(e.onlineMeetingProvider))) {
    return "Teams";
  }
  const room = e.location?.displayName?.trim();
  return room || (e.isOnlineMeeting ? "Online" : "—");
}

/** "30m", "1h 15m", "2h" */
export function formatDuration(ms: number): string {
  const totalMin = Math.max(0, Math.round(ms / 60_000));
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  if (h === 0) return `${m}m`;
  if (m === 0) return `${h}h`;
  return `${h}h ${m}m`;
}

/** "in 12m", "in 1h 5m", "in 3h", "now" */
export function formatCountdown(ms: number): string {
  if (ms <= 30_000) return "now";
  const totalMin = Math.round(ms / 60_000);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  if (h === 0) return `in ${m}m`;
  if (m === 0) return `in ${h}h`;
  return `in ${h}h ${m}m`;
}

function formatClock(d: Date, tz: string): string {
  return d.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit", timeZone: tz });
}

/** Build the widget lines. First (next) meeting shows length + countdown; second shows start time + countdown.
 *  When the next meeting starts within URGENT_MS, its countdown is passed through `colorize` (e.g. red).
 *  `colorize` defaults to plain text so the function stays UI-free. Clocks render in `tz`. */
export const URGENT_MS = 5 * 60_000;

export function buildWidgetLines(
  events: GraphEvent[],
  now: Date,
  count = 2,
  colorize: (text: string, urgent: boolean) => string = (t) => t,
  tz = "UTC",
): string[] {
  const next = pickNextEvents(events, now, count, tz);
  return next.map((e, i) => {
    const bullet = i === 0 ? "●" : "○";
    const running = e.startsAt.getTime() <= now.getTime();
    const msToStart = e.startsAt.getTime() - now.getTime();
    const when = running
      ? `ends ${formatCountdown(e.endsAt.getTime() - now.getTime())}`
      : formatCountdown(msToStart);
    const urgent = !running && msToStart <= URGENT_MS;
    const whenText = colorize(when, urgent);
    if (i === 0) {
      return `${bullet} ${e.subject} — ${e.location} · ${formatDuration(e.durationMs)} · ${whenText}`;
    }
    return `${bullet} ${e.subject} — ${e.location} · ${formatClock(e.startsAt, tz)} · ${whenText}`;
  });
}

/**
 * Compact single-line summary for pi's footer status area, e.g.
 * "● Arch review (Lille Sal) 1h 30m ends in 40m · ○ 1:1 (Teams) in 1h 50m".
 * Returns empty string when there is nothing left today.
 */
export function buildStatusLine(events: GraphEvent[], now: Date, count = 2, tz = "UTC"): string {
  const next = pickNextEvents(events, now, count, tz);
  if (next.length === 0) return "";
  return next
    .map((e) => {
      const running = e.startsAt.getTime() <= now.getTime();
      const when = running
        ? `ends ${formatCountdown(e.endsAt.getTime() - now.getTime())}`
        : formatCountdown(e.startsAt.getTime() - now.getTime());
      const title = e.subject.length > 24 ? `${e.subject.slice(0, 23)}…` : e.subject;
      return `● ${title} (${e.location}) ${when}`;
    })
    .join("  ·  ");
}

// ---------------------------------------------------------------------------
// Live Graph API access (client credentials flow)
// ---------------------------------------------------------------------------

export interface OAuthConfig {
  tenantId: string;
  clientId: string;
  /**
   * Inside an sbx sandbox this env var holds a placeholder (sentinel), never
   * the real secret: the literal "proxy-managed" (engine-injected credentials)
   * or an "sbx-cs-…"-prefixed value (custom secrets via
   * `sbx secret set-custom --env O365_CLIENT_SECRET …`). The egress proxy
   * swaps the placeholder for the real secret on the AAD token request.
   * On the host it is the real secret and goes in the request body
   * (client_secret_post).
   */
  clientSecret: string;
  upn: string; // calendar owner, e.g. ska@vertica.dk
  /** IANA zone the user lives in; used for "today" filtering and clocks. */
  tz: string;
}

/** Sentinel values sbx sets for proxy-managed credentials inside the container. */
export const PROXY_SENTINELS = ["proxy-managed", "sbx-cs-"];

/** True when the client secret is an sbx placeholder, not the real secret. */
export function isProxySentinel(clientSecret: string): boolean {
  return PROXY_SENTINELS.some((s) => clientSecret === s || clientSecret.startsWith(s));
}

/**
 * "basic-proxy": secret is an sbx sentinel — send it via the Authorization
 * header (client_secret_basic, RFC 6749 §2.3.1) so the egress proxy can swap
 * it. "body": real secret (e.g. on the host) — send it in the POST body
 * (client_secret_post).
 */
export type AuthMode = "basic-proxy" | "body";

export function authMode(cfg: OAuthConfig): AuthMode {
  return isProxySentinel(cfg.clientSecret) ? "basic-proxy" : "body";
}

export function oauthConfigFromEnv(env: NodeJS.ProcessEnv = process.env): OAuthConfig | undefined {
  const tenantId = env.O365_TENANT_ID;
  const clientId = env.O365_CLIENT_ID;
  const clientSecret = env.O365_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) return undefined;
  // O365_TZ is optional; default matches the UPN default (CET/CEST). The
  // sandbox VM runs UTC, so this must NOT come from the machine's zone.
  return {
    tenantId,
    clientId,
    clientSecret,
    upn: env.O365_UPN || "ska@vertica.dk",
    tz: env.O365_TZ || "Europe/Copenhagen",
  };
}

export interface TokenRequest {
  url: string;
  headers: Record<string, string>;
  body: string;
}

/** Build the AAD client-credentials token request for the configured mode. */
export function buildTokenRequest(cfg: OAuthConfig): TokenRequest {
  const url = `https://login.microsoftonline.com/${cfg.tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });
  const headers: Record<string, string> = {
    "Content-Type": "application/x-www-form-urlencoded",
  };
  if (authMode(cfg) === "basic-proxy") {
    // client_secret_basic. The password is the sbx placeholder; the sandbox
    // egress proxy swaps it for the real client secret. The body carries no
    // secret.
    headers["Authorization"] =
      "Basic " + Buffer.from(`${cfg.clientId}:${cfg.clientSecret}`).toString("base64");
  } else {
    params.set("client_id", cfg.clientId);
    params.set("client_secret", cfg.clientSecret);
  }
  return { url, headers, body: params.toString() };
}

interface CachedToken {
  token: string;
  expiresAtMs: number;
}

let cachedToken: CachedToken | undefined;

/** Test helper: drop the cached access token. */
export function clearTokenCache(): void {
  cachedToken = undefined;
}

export async function getAccessToken(cfg: OAuthConfig, signal?: AbortSignal): Promise<string> {
  // Tokens live ~1h; reuse until 5 min before expiry.
  const now = Date.now();
  if (cachedToken && cachedToken.expiresAtMs - now > 5 * 60_000) {
    return cachedToken.token;
  }
  const req = buildTokenRequest(cfg);
  const res = await fetch(req.url, { method: "POST", headers: req.headers, body: req.body, signal });
  if (!res.ok) {
    throw new Error(`OAuth token request failed: HTTP ${res.status}`);
  }
  const data = (await res.json()) as { access_token?: string; expires_in?: number };
  if (!data.access_token) throw new Error("OAuth token response missing access_token");
  cachedToken = {
    token: data.access_token,
    expiresAtMs: now + (data.expires_in ?? 3600) * 1000,
  };
  return cachedToken.token;
}

export interface GraphQuery {
  /** Inclusive range start (ISO, any zone — sent as UTC). */
  start: Date;
  /** Exclusive range end (ISO). */
  end: Date;
  upn: string;
}

/**
 * [start, end) of "today" as seen from `tz`: midnight local (DST-safe) to
 * midnight local next day, as real instants. The VM may run UTC, so the
 * window must be derived from the zone, not the machine clock.
 */
export function buildDayWindow(now: Date, tz: string): { start: Date; end: Date } {
  const ymd = wallDateInZone(now, tz);
  const [y, m, d] = ymd.split("-").map(Number);
  const start = parseGraphDateTime(`${ymd}T00:00:00.000`, tz)!;
  const nextYmd = wallDateInZone(new Date(Date.UTC(y, m - 1, d + 1, 12)), tz);
  const end = parseGraphDateTime(`${nextYmd}T00:00:00.000`, tz)!;
  return { start, end };
}

/**
 * calendarView expands recurring series into instances (plain /events does not),
 * which is what we need for "next two meetings".
 */
export function buildCalendarViewUrl(q: GraphQuery): string {
  const params = new URLSearchParams({
    startDateTime: q.start.toISOString(),
    endDateTime: q.end.toISOString(),
    $orderby: "start/dateTime",
    $select: "subject,start,end,location,isOnlineMeeting,onlineMeetingProvider",
    $top: "20",
  });
  return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(q.upn)}/calendarview?${params}`;
}

export async function fetchTodayEvents(
  cfg: OAuthConfig,
  now: Date,
  signal?: AbortSignal,
): Promise<GraphEvent[]> {
  const token = await getAccessToken(cfg, signal);
  const { start: dayStart, end: dayEnd } = buildDayWindow(now, cfg.tz);

  const url = buildCalendarViewUrl({ start: dayStart, end: dayEnd, upn: cfg.upn });
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      // Return dateTime values in UTC (with Z suffix) so wall-clock parsing
      // cannot silently fall back to the VM's zone.
      Prefer: 'outlook.timezone="UTC"',
    },
    signal,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph calendarview failed: HTTP ${res.status} ${text.slice(0, 200)}`);
  }
  const data = (await res.json()) as { value?: GraphEvent[] };
  return data.value ?? [];
}
