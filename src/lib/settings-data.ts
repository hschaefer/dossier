import "server-only";

import type { Prisma } from "@/generated/prisma/client";
import { ActionError } from "@/lib/action";
import { agentSearchInputSchema } from "@/lib/agent/search";
import { lockAgentSearches } from "@/lib/agent/searches";
import { EMPTY_QUERY, searchQuerySchema } from "@/lib/job-search/search";
import { prisma } from "@/lib/prisma";
import { DEFAULT_BASE_THEME, settingsFromRows, settingValue, type SettingsData } from "@/lib/settings";

/**
 * The settings in effect. Nothing in the database ties the base theme
 * setting to its theme, so a deleted custom theme reads as Editorial here.
 */
export async function loadSettings(): Promise<SettingsData> {
  const settings = settingsFromRows(await prisma.setting.findMany());
  // Compatibility for the single-search screen/executor until they use the
  // search service directly. Never migrate old settings at read time.
  const search = await prisma.agentSearch.findFirst({ where: { enabled: true }, orderBy: [{ sort: "asc" }, { createdAt: "asc" }, { id: "asc" }] });
  settings.agentQuery = search ? searchQuerySchema.parse(search.query) : EMPTY_QUERY;
  if (settings.baseThemeId !== DEFAULT_BASE_THEME) {
    const theme = await prisma.theme.findUnique({ where: { id: settings.baseThemeId }, select: { id: true } });
    if (!theme) return { ...settings, baseThemeId: DEFAULT_BASE_THEME };
  }
  return settings;
}

/**
 * Saves the given settings together; null removes one, so it reads as not
 * set. Pass a transaction client to save them as part of that transaction.
 */
export async function saveSettings(changes: Partial<SettingsData>, db?: Prisma.TransactionClient): Promise<void> {
  if (!db) return prisma.$transaction((tx) => saveSettings(changes, tx));
  // Lock before any setting writes, matching import/wipe lock ordering.
  if (changes.agentQuery !== undefined) {
    await lockAgentSearches(db);
    if (await db.agentSearch.count() > 1) throw new ActionError("This single-search form cannot edit multiple Agent searches.");
  }
  for (const [key, change] of Object.entries(changes)) {
    if (change === undefined) continue;
    if (key === "agentQuery") {
      // The legacy form can edit a single search, not a collection.
      const first = await db.agentSearch.findFirst({ orderBy: [{ sort: "asc" }, { createdAt: "asc" }, { id: "asc" }] });
      const query = searchQuerySchema.parse(change);
      if (!query.keywords) {
        if (first) await db.agentSearch.delete({ where: { id: first.id } });
      } else {
        const data = agentSearchInputSchema.parse({ name: first?.name ?? "My search", query, enabled: first?.enabled ?? true, sort: first?.sort ?? 0 });
        if (first) await db.agentSearch.update({ where: { id: first.id }, data });
        else await db.agentSearch.create({ data });
      }
      await db.setting.deleteMany({ where: { key: "agentQuery" } });
      continue;
    }
    const value = settingValue(change);
    if (value === null) await db.setting.deleteMany({ where: { key } });
    else await db.setting.upsert({ where: { key }, create: { key, value }, update: { value } });
  }
}
