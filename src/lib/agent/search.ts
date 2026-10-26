import { z } from "zod";

import { EMPTY_QUERY, searchQuerySchema, type SearchQuery } from "@/lib/job-search/search";

/** Fixed product and source-budget cap, shared by forms and server writes. */
export const MAX_AGENT_SEARCHES = 3;
export const agentSearchInputSchema = z.object({
  name: z.string().trim().min(1, "Give the search a name.").max(100),
  query: searchQuerySchema.refine((query) => Boolean(query.keywords), "Add what to look for."),
  enabled: z.boolean().default(true),
  sort: z.number().int().min(0).default(0),
});
export type AgentSearchInput = z.input<typeof agentSearchInputSchema>;
export type SavedAgentSearch = { id: string; name: string; query: SearchQuery; enabled: boolean; sort: number; createdAt: Date; updatedAt: Date };

/** Same recovery rules as the SQL upgrade: keep meaningful keywords, repair
 * optional fields and truncate the place rather than lose the whole search. */
export function legacyAgentQuery(value: string): SearchQuery | null {
  let raw: unknown;
  try { raw = JSON.parse(value); } catch { return null; }
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const q = raw as Record<string, unknown>;
  const keywords = searchQuerySchema.shape.keywords.safeParse(q.keywords);
  if (!keywords.success || !keywords.data) return null;
  return {
    keywords: keywords.data,
    place: typeof q.place === "string" ? Array.from(q.place.trim()).slice(0, 200).join("").trim() : "",
    country: searchQuerySchema.shape.country.catch(EMPTY_QUERY.country).parse(q.country),
    radiusKm: searchQuerySchema.shape.radiusKm.catch(EMPTY_QUERY.radiusKm).parse(q.radiusKm),
    remoteOnly: searchQuerySchema.shape.remoteOnly.catch(EMPTY_QUERY.remoteOnly).parse(q.remoteOnly),
    postedWithin: searchQuerySchema.shape.postedWithin.catch(EMPTY_QUERY.postedWithin).parse(q.postedWithin),
  };
}
