/**
 * Zeldoc.ai web search for Pi.
 *
 * Tools:
 * - web_search: query Zeldoc.ai's private search (POST /v1/search/zeldoc-search)
 * - fetch_content: download a URL locally and return its readable text
 *
 * The API key comes from Pi's credential store (/login, Zeldoc.ai) or the
 * ZELDOC_API_KEY environment variable. Nothing else leaves your machine.
 *
 * Vendored from https://docs.zeldoc.ai/search/pi/zeldoc_search.ts
 * (https://docs.zeldoc.ai/web-search/)
 */

import { StringEnum } from "@earendil-works/pi-ai";
import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import { Type } from "typebox";

const SEARCH_URL = "https://api.zeldoc.ai/v1/search/zeldoc-search";
const MAX_RESULTS = 10;

interface SearchResult {
  title: string;
  url: string;
  snippet: string;
  date?: string;
}

async function resolveApiKey(ctx: ExtensionContext): Promise<string | undefined> {
  const stored = await ctx.modelRegistry
    .getApiKeyForProvider("zeldoc")
    .catch(() => undefined);
  return stored || process.env.ZELDOC_API_KEY || undefined;
}

async function zeldocSearch(
  apiKey: string,
  body: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<SearchResult[]> {
  const response = await fetch(SEARCH_URL, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
    signal,
  });
  if (!response.ok) {
    throw new Error(`Zeldoc.ai search failed: HTTP ${response.status}`);
  }
  const data = (await response.json()) as { results?: SearchResult[] };
  return data.results ?? [];
}

function formatResults(results: SearchResult[]): string {
  if (results.length === 0) return "No results found.";
  return results
    .slice(0, MAX_RESULTS)
    .map((r, i) =>
      [
        `--- Result ${i + 1} ---`,
        `Title: ${r.title}`,
        `Link: ${r.url}`,
        r.date ? `Date: ${r.date}` : undefined,
        `Snippet: ${r.snippet}`,
      ]
        .filter(Boolean)
        .join("\n"),
    )
    .join("\n\n");
}

/** Crude but dependency-free HTML to text. */
function htmlToText(html: string): string {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, "")
    .replace(/<\/(p|div|br|li|h[1-6]|tr|section|article)>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function fetchContent(url: string, signal?: AbortSignal): Promise<string> {
  const response = await fetch(url, {
    headers: {
      Accept: "text/html,text/plain,text/markdown;q=0.9,*/*;q=0.5",
      "User-Agent": "pi-zeldoc-search/1.0",
    },
    signal,
  });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}`);
  }
  const text = await response.text();
  const type = response.headers.get("content-type") ?? "";
  return type.includes("html") ? htmlToText(text) : text;
}

const SEARCH_PARAMS = Type.Object({
  query: Type.String({ description: "Search query" }),
  time_range: Type.Optional(
    StringEnum(["day", "week", "month", "year"] as const, {
      description: "Only results from this period",
    }),
  ),
  language: Type.Optional(
    Type.String({ description: "Two-letter language code, e.g. en, de, da" }),
  ),
  engines: Type.Optional(
    Type.String({
      description: "Comma-separated engines to use, e.g. github,stackoverflow",
    }),
  ),
});

const FETCH_PARAMS = Type.Object({
  url: Type.String({ description: "URL to fetch" }),
  maxChars: Type.Optional(
    Type.Number({
      description: "Max characters to return (default 20000)",
      minimum: 500,
      maximum: 200_000,
    }),
  ),
});

export default function (pi: ExtensionAPI): void {
  pi.registerTool({
    name: "web_search",
    label: "Web Search",
    description:
      "Search the web through Zeldoc.ai's private search. Returns titles, " +
      "URLs and snippets. Call fetch_content on a result when the snippet " +
      "is not enough.",
    promptSnippet: "Search the web (private, via Zeldoc.ai)",
    parameters: SEARCH_PARAMS,
    async execute(_toolCallId, params, signal, _onUpdate, ctx) {
      const apiKey = await resolveApiKey(ctx);
      if (!apiKey) {
        throw new Error(
          "No Zeldoc.ai API key. Run /login and pick Zeldoc.ai, or set ZELDOC_API_KEY.",
        );
      }
      const results = await zeldocSearch(apiKey, params, signal);
      return {
        content: [{ type: "text", text: formatResults(results) }],
        details: { query: params.query, resultCount: results.length },
      };
    },
  });

  pi.registerTool({
    name: "fetch_content",
    label: "Fetch Content",
    description:
      "Download a URL on this machine and return its readable text. " +
      "Use it to read a page found with web_search.",
    promptSnippet: "Fetch a web page as text",
    parameters: FETCH_PARAMS,
    async execute(_toolCallId, params, signal) {
      const content = await fetchContent(params.url, signal);
      const maxChars = params.maxChars ?? 20_000;
      const text =
        content.length > maxChars
          ? `${content.slice(0, maxChars)}\n\n[truncated at ${maxChars} characters]`
          : content;
      return {
        content: [{ type: "text", text }],
        details: { url: params.url, length: content.length },
      };
    },
  });
}
