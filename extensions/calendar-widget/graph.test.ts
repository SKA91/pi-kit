/**
 * Smoke test for extensions/calendar-widget/graph.ts
 *
 * Run: node --experimental-strip-types extensions/calendar-widget/graph.test.ts
 */

import assert from "node:assert";
import {
  authMode,
  buildDayWindow,
  buildStatusLine,
  buildTokenRequest,
  buildWidgetLines,
  formatCountdown,
  formatDuration,
  oauthConfigFromEnv,
  parseGraphDateTime,
  pickNextEvents,
  PROXY_SENTINELS,
  timeZoneOffsetMinutes,
  URGENT_MS,
  wallDateInZone,
} from "./graph.ts";

// Use a fixed "now" so output is deterministic regardless of the host's
// zone: Wed 2026-02-04 11:20 UTC (an explicit instant, not host-local).
const now = new Date(Date.UTC(2026, 1, 4, 11, 20, 0));

const ev = (
  subject: string,
  startH: number,
  startM: number,
  endH: number,
  endM: number,
  opts: { room?: string; online?: boolean; dayOffset?: number; tz?: string } = {},
) => {
  const day = (h: number, m: number, offset = 0) => {
    const d = new Date(now);
    d.setDate(d.getDate() + offset);
    d.setHours(h, m, 0, 0);
    return d;
  };
  const iso = (d: Date) => {
    const pad = (n: number) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:00.000`;
  };
  return {
    subject,
    start: { dateTime: iso(day(startH, startM, opts.dayOffset)), timeZone: opts.tz ?? null },
    end: { dateTime: iso(day(endH, endM, opts.dayOffset)), timeZone: opts.tz ?? null },
    location: { displayName: opts.room ?? null },
    isOnlineMeeting: opts.online ?? false,
    onlineMeetingProvider: opts.online ? "teamsForBusiness" : null,
  };
};

const events = [
  ev("Yesterday retro", 15, 0, 16, 0, { room: "Lille Sal", dayOffset: -1 }),
  ev("Morning sync", 9, 0, 9, 30, { online: true }), // already over at 11:20
  ev("Architecture review", 10, 30, 12, 0, { room: "Lille Sal / 2.1" }), // in progress
  ev("Teams check-in", 12, 30, 13, 0, { online: true }),
  ev("Customer workshop", 14, 30, 16, 30, { room: "Store Sal / 2.2" }),
  ev("Tomorrow 1:1", 9, 0, 9, 30, { online: true, dayOffset: 1 }),
];

// --- parseGraphDateTime ---
assert.ok(parseGraphDateTime("2026-02-04T10:30:00.000"));
assert.equal(parseGraphDateTime("garbage"), null);
assert.equal(parseGraphDateTime(undefined), null);
// Graph's tick precision: 7 fractional digits must not fail parsing (regression).
const ticks = parseGraphDateTime("2026-02-04T10:30:00.0000000")!;
assert.equal(ticks.toISOString(), "2026-02-04T10:30:00.000Z");

// --- timezone handling (the VM runs UTC; the user lives in Copenhagen) ---
// Naive string + explicit zone: 14:00 in Copenhagen (UTC+1 in February) is 13:00 UTC.
const cph = parseGraphDateTime("2026-02-04T14:00:00.000", "Europe/Copenhagen")!;
assert.equal(cph.toISOString(), "2026-02-04T13:00:00.000Z");
// Same string, no zone: falls back to UTC (VM zone) — the old buggy behavior.
const utc = parseGraphDateTime("2026-02-04T14:00:00.000", "UTC")!;
assert.equal(utc.toISOString(), "2026-02-04T14:00:00.000Z");
// Explicit offsets/Z pass through untouched.
assert.equal(parseGraphDateTime("2026-02-04T14:00:00Z", "Europe/Copenhagen")!.toISOString(), "2026-02-04T14:00:00.000Z");
assert.equal(parseGraphDateTime("2026-02-04T14:00:00+05:30", "UTC")!.toISOString(), "2026-02-04T08:30:00.000Z");
// Windows tz names are not IANA — parsing falls back to null for naive strings.
assert.equal(parseGraphDateTime("2026-02-04T14:00:00.000", "W. Europe Standard Time"), null);
// Offset helper sanity: Copenhagen is UTC+1 in February, UTC+2 in July.
assert.equal(timeZoneOffsetMinutes("Europe/Copenhagen", new Date(2026, 1, 4)), 60);
assert.equal(timeZoneOffsetMinutes("Europe/Copenhagen", new Date(2026, 6, 4)), 120);
assert.equal(timeZoneOffsetMinutes("Not/AZone", new Date()), null);
// Wall-date in zone: 23:30 UTC on Feb 4 is already Feb 5 in Copenhagen.
assert.equal(wallDateInZone(new Date(Date.UTC(2026, 1, 4, 23, 30)), "Europe/Copenhagen"), "2026-02-05");
assert.equal(wallDateInZone(new Date(Date.UTC(2026, 1, 4, 23, 30)), "UTC"), "2026-02-04");

// --- pickNextEvents: filters past/tomorrow, sorts, takes 2 ---
// Events carry no timeZone → naive strings are read in the VM's zone (UTC);
// `now` was built in the same zone, so the selection is unchanged.
const next = pickNextEvents(events, now, 2, "UTC");
assert.deepEqual(
  next.map((e) => e.subject),
  ["Architecture review", "Teams check-in"],
);
assert.equal(next[0].location, "Lille Sal / 2.1");
assert.equal(next[1].location, "Teams");
assert.equal(next[0].durationMs, 90 * 60_000);

// --- pickNextEvents honors event timeZone over the fallback zone ---
// 12:30-13:00 "W. Europe Standard Time" is a Windows name → unparseable →
// falls back to the user zone; but a real IANA zone must win over the
// fallback: a 00:30-01:00 Copenhagen event (Feb 4) is 23:30 UTC Feb 3, i.e.
// YESTERDAY's wall date in UTC — only visible when the zone is applied.
const cphEvents = [ev("Night owl", 0, 30, 1, 0, { online: true, tz: "Europe/Copenhagen" })];
assert.deepEqual(pickNextEvents(cphEvents, now, 2, "UTC").map((e) => e.subject), []);

// --- formatters ---
assert.equal(formatDuration(90 * 60_000), "1h 30m");
assert.equal(formatDuration(45 * 60_000), "45m");
assert.equal(formatDuration(2 * 3_600_000), "2h");
assert.equal(formatCountdown(20_000), "now");
assert.equal(formatCountdown(12 * 60_000), "in 12m");
assert.equal(formatCountdown(65 * 60_000), "in 1h 5m");
assert.equal(formatCountdown(3 * 3_600_000), "in 3h");

// --- widget lines: next meeting shows duration, second shows start time ---
const lines = buildWidgetLines(events, now, 2, (t) => t, "UTC");
assert.equal(lines.length, 2);
assert.match(lines[0], /● Architecture review — Lille Sal \/ 2\.1 · 1h 30m · ends in 40m/);
assert.match(lines[1], /○ Teams check-in — Teams · 12:30 · in 1h 10m/);

// --- clocks render in the user's zone, not the VM's ---
// "First" is 13:00-14:00 Copenhagen (in 40m from now), "Zoned" is 12:30-13:00
// UTC (= 13:30 Copenhagen, in 1h 10m). Zoned lands on line 2, whose clock
// must show 13:30 (Copenhagen), not 12:30 (the VM's UTC).
const tzLines = buildWidgetLines(
  [
    ev("First", 13, 0, 14, 0, { online: true, tz: "Europe/Copenhagen" }),
    ev("Zoned", 12, 30, 13, 0, { online: true, tz: "UTC" }),
  ],
  now,
  2,
  (t) => t,
  "Europe/Copenhagen",
);
assert.match(tzLines[0], /in 40m/);
assert.match(tzLines[1], /13:30/);

// --- urgent countdown (≤ 5 min to next start) gets flagged to colorize ---
const soon = new Date(now);
soon.setMinutes(soon.getMinutes() + 3);
const urgentEvents = [
  ev("Daily standup", 11, 23, 11, 38, { online: true }), // starts in 3 min
  ev("Later meeting", 15, 0, 16, 0, { room: "Lille Sal" }),
];
const urgentLines = buildWidgetLines(urgentEvents, now, 2, (t, urgent) =>
  urgent ? `RED(${t})` : t,
  "UTC",
);
assert.match(urgentLines[0], /RED\(in 3m\)/);
assert.match(urgentLines[1], /in 3h 40m(?!.*RED)/);

// --- boundary: exactly 5 min is urgent, 6 min is not ---
const edge = buildWidgetLines([ev("Edge", 11, 25, 12, 0)], now, 1, (_t, urgent) => urgent ? "URGENT" : "CALM", "UTC");
assert.match(edge[0], /URGENT/);
const edge2 = buildWidgetLines([ev("Edge", 11, 26, 12, 0)], now, 1, (_t, urgent) => urgent ? "URGENT" : "CALM", "UTC");
assert.match(edge2[0], /CALM/);

// --- empty result: nothing left today ---
assert.deepEqual(buildWidgetLines([events[0], events[5]], now, 2, (t) => t, "UTC"), []);

// --- day window: "today" in Copenhagen starts 23:00 UTC (February) ---
const febNoonUtc = new Date(Date.UTC(2026, 1, 4, 11, 0, 0)); // 12:00 Copenhagen
const win = buildDayWindow(febNoonUtc, "Europe/Copenhagen");
assert.equal(win.start.toISOString(), "2026-02-03T23:00:00.000Z");
assert.equal(win.end.toISOString(), "2026-02-04T23:00:00.000Z");

// --- env config ---
assert.equal(oauthConfigFromEnv({}), undefined);
assert.deepEqual(
  oauthConfigFromEnv({
    O365_TENANT_ID: "t",
    O365_CLIENT_ID: "c",
    O365_CLIENT_SECRET: "s",
  }),
  { tenantId: "t", clientId: "c", clientSecret: "s", upn: "ska@vertica.dk", tz: "Europe/Copenhagen" },
);
assert.equal(
  oauthConfigFromEnv({
    O365_TENANT_ID: "t",
    O365_CLIENT_ID: "c",
    O365_CLIENT_SECRET: "s",
    O365_TZ: "Europe/Berlin",
  })!.tz,
  "Europe/Berlin",
);

// --- auth modes: host secret goes in the body, sbx sentinel goes in the header ---
const hostCfg = oauthConfigFromEnv({ O365_TENANT_ID: "t", O365_CLIENT_ID: "c", O365_CLIENT_SECRET: "s" })!;
assert.equal(authMode(hostCfg), "body");
const hostReq = buildTokenRequest(hostCfg);
assert.ok(hostReq.url.startsWith("https://login.microsoftonline.com/t/oauth2/v2.0/token"));
assert.ok(!hostReq.headers["Authorization"], "host mode must not set Authorization");
assert.match(hostReq.body, /client_secret=s/);
assert.match(hostReq.body, /grant_type=client_credentials/);

const proxyCfg = { ...hostCfg, clientSecret: PROXY_SENTINELS[0] };
assert.equal(authMode(proxyCfg), "basic-proxy");
const proxyReq = buildTokenRequest(proxyCfg);
const expectedBasic = Buffer.from(`c:${PROXY_SENTINELS[0]}`).toString("base64");
assert.equal(proxyReq.headers["Authorization"], `Basic ${expectedBasic}`);
assert.ok(!/client_secret/.test(proxyReq.body), "proxy mode must not leak the sentinel in the body");
assert.match(proxyReq.body, /grant_type=client_credentials/);

// --- sbx custom-secret placeholder ("sbx-cs-…") also switches to basic auth ---
const sbxCfg = { ...hostCfg, clientSecret: "sbx-cs-GYgOlM3mcdCfs1gm" };
assert.equal(authMode(sbxCfg), "basic-proxy");
const sbxReq = buildTokenRequest(sbxCfg);
const sbxBasic = Buffer.from(`c:${sbxCfg.clientSecret}`).toString("base64");
assert.equal(sbxReq.headers["Authorization"], `Basic ${sbxBasic}`);
assert.ok(!/client_secret/.test(sbxReq.body), "sbx placeholder must not go in the body");

console.log("All calendar-widget tests passed ✓");
console.log("Example widget render:");
for (const line of lines) console.log("  " + line);
