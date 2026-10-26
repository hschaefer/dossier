import { readFile } from "node:fs/promises";

import { Client } from "pg";
import { describe, expect, it } from "vitest";

import { testDatabaseUrl } from "@/test/db";

import { legacyAgentQuery } from "./search";

// Replay the actual migration against a minimal pre-upgrade schema. Each
// transaction rolls back, including its schema, without touching public data.
describe("Agent search SQL upgrade", () => {
  it.each([
    [JSON.stringify({ keywords: "  Designer ", place: " Berlin ", country: " DE ", radiusKm: 50, remoteOnly: true, postedWithin: "week" }), true],
    [JSON.stringify({ keywords: "Engineer" }), true],
    [JSON.stringify({ keywords: "Engineer", place: " x".repeat(300), country: "invalid", radiusKm: 42, remoteOnly: "yes", postedWithin: "invalid" }), true],
    [JSON.stringify({ keywords: " " }), false],
    [JSON.stringify({ keywords: "\t\n " }), false],
    [JSON.stringify({ keywords: 42 }), false],
    ["{broken", false],
    ["null", false],
    [null, false],
  ])("migrates only a meaningful query (%s)", async (value, expected) => {
    const db = new Client({ connectionString: testDatabaseUrl() });
    await db.connect();
    try {
      await db.query("BEGIN");
      await db.query('CREATE SCHEMA agent_search_upgrade_test');
      await db.query('SET LOCAL search_path TO agent_search_upgrade_test');
      await db.query('CREATE TABLE "Setting" ("key" text PRIMARY KEY, "value" text)');
      await db.query('CREATE TABLE "AgentCandidate" ("id" text PRIMARY KEY)');
      if (value !== null) await db.query('INSERT INTO "Setting" VALUES ($1, $2)', ["agentQuery", value]);
      await db.query('INSERT INTO "Setting" VALUES ($1, $2)', ["jsearchKeyEnc", "untouched"]);
      await db.query(await readFile("prisma/migrations/20261026100000_agent_searches/migration.sql", "utf8"));
      const searches = await db.query('SELECT * FROM "AgentSearch"');
      expect(searches.rows).toHaveLength(expected ? 1 : 0);
      if (expected) {
        expect(searches.rows[0]).toMatchObject({ id: "agent-search-legacy", name: "My search", enabled: true, sort: 0 });
        expect(searches.rows[0].query.keywords).toBe(JSON.parse(value!).keywords.trim());
        expect(searches.rows[0].query.country).toBe(value?.includes("DE") ? "de" : "");
        expect(searches.rows[0].query.place.length).toBeLessThanOrEqual(200);
        expect(searches.rows[0].query).toEqual(legacyAgentQuery(value!));
      }
      expect((await db.query('SELECT * FROM "Setting"')).rows).toEqual([{ key: "jsearchKeyEnc", value: "untouched" }]);
      expect((await db.query('SELECT * FROM "AgentCandidateSearch"')).rows).toEqual([]);
    } finally {
      await db.query("ROLLBACK");
      await db.end();
    }
  });
});
