"use client";

import Link from "next/link";
import { useState } from "react";

import { useTrackSave } from "@/components/save-tracker";
import { toastResult } from "@/components/toast";
import { Card } from "@/components/ui/card";
import { Select } from "@/components/ui/select";
import { Toggle } from "@/components/ui/toggle";
import type { SavedAgentSearch } from "@/lib/agent/search";
import { AGENT_MAX_SOURCE_REQUESTS_PER_DAY } from "@/lib/agent/limits";
import type { Country } from "@/lib/job-search/countries";
import { type JobSource, type SearchQuery } from "@/lib/job-search/search";

import { saveAgentSettings } from "../agent/actions";
import { RecurringSearches } from "./recurring-searches";
import { SettingsSectionIntro } from "./settings-screen";

export type AgentSettings = {
  enabled: boolean;
  hour: number;
  minScore: number;
  maxScores: number;
  maxDrafts: number;
  autoDraft: boolean;
};

const HOURS = Array.from({ length: 24 }, (_, hour) => hour);
const MIN_SCORES = [50, 55, 60, 65, 70, 75, 80, 85, 90];
const SCORE_COUNTS = Array.from({ length: 30 }, (_, index) => index + 1);
const DRAFT_COUNTS = Array.from({ length: 11 }, (_, index) => index);
const clock = (hour: number) => `${String(hour).padStart(2, "0")}:00`;
const plural = (count: number, one: string, many: string) => `${count} ${count === 1 ? one : many}`;

/**
 * Shared daily schedule and AI limits, with explicit saved-search management.
 */
export function AgentSection({
  initial,
  source,
  aiReady,
  countries,
  timeZone,
  searches,
  suggestedQuery,
}: {
  initial: AgentSettings;
  /** The source in Settings → Job search; null while job search is off. */
  source: JobSource | null;
  aiReady: boolean;
  countries: readonly Country[];
  /** The user's time zone, which the hour is in. */
  timeZone: string;
  searches: SavedAgentSearch[];
  suggestedQuery?: SearchQuery;
}) {
  const track = useTrackSave();
  const [settings, setSettings] = useState(initial);

  const change = (next: Partial<AgentSettings>) => {
    const previous = settings;
    setSettings({ ...settings, ...next });
    void track(saveAgentSettings(next)).then((result) => {
      if (!toastResult(result).ok) setSettings(previous);
    });
  };

  return (
    <>
      <SettingsSectionIntro title="Job agent">
        Once a day it runs all enabled searches, scores the new postings against your profile, and keeps the best on the Agent screen. It never sends or applies.
      </SettingsSectionIntro>

      {source === "ARBEITSAGENTUR" ? (
        // It searches on its own, which the BA's terms forbid most clearly (#223).
        <p role="note" className="rounded-card bg-warning-surface px-4 py-3 text-meta text-warning">
          The agent doesn&apos;t search the Bundesagentur für Arbeit. Choose JSearch or Adzuna in{" "}
          <Link href="/settings?section=jobs" className="font-medium text-warning underline underline-offset-2">
            Settings → Job search
          </Link>{" "}
          for it. Manual Find jobs searches are unaffected.
        </p>
      ) : (!source || !aiReady) && (
        <p role="note" className="rounded-card bg-warning-surface px-4 py-3 text-meta text-warning">
          {!source ? "The agent searches the source you choose in " : "The agent scores postings with the AI provider in "}
          <Link href={!source ? "/settings?section=jobs" : "/settings?section=ai"} className="font-medium text-warning underline underline-offset-2">
            {!source ? "Settings → Job search" : "Settings → AI provider"}
          </Link>
          {!source ? ", which is off. Choose one first." : ", which isn't set up yet. Choose one first."}
        </p>
      )}

      <Card padding="none" className="px-7 py-4">
        <Toggle
          label="Run the job agent"
          description={
            settings.enabled
              ? `On. It runs once a day after ${clock(settings.hour)} your time (${timeZone}), and you can run it now and then on the Agent screen.`
              : "Off. Nothing is searched or scored until it's on."
          }
          checked={settings.enabled}
          onCheckedChange={(enabled) => change({ enabled })}
        />
      </Card>

      <RecurringSearches initial={searches} suggestedQuery={suggestedQuery} source={source} countries={countries} />

      <Card padding="none" className="grid gap-x-[18px] gap-y-5 px-7 py-[26px] sm:grid-cols-2">
        <Select label="Runs after" hint={`Your time (${timeZone})`} value={String(settings.hour)} onChange={(event) => change({ hour: Number(event.target.value) })}>
          {HOURS.map((hour) => (
            <option key={hour} value={hour}>
              {clock(hour)}
            </option>
          ))}
        </Select>
        <Select
          label="Minimum score"
          hint="Postings scored lower stay out of the queue."
          value={String(settings.minScore)}
          onChange={(event) => change({ minScore: Number(event.target.value) })}
        >
          {[...new Set([...MIN_SCORES, settings.minScore])].sort((a, b) => a - b).map((score) => (
            <option key={score} value={score}>
              {score} of 100
            </option>
          ))}
        </Select>
        <Select label="Postings scored per run" value={String(settings.maxScores)} onChange={(event) => change({ maxScores: Number(event.target.value) })}>
          {SCORE_COUNTS.map((count) => (
            <option key={count} value={count}>
              {count}
            </option>
          ))}
        </Select>
        <Select
          label="Drafts per run, at most"
          value={String(settings.maxDrafts)}
          disabled={!settings.autoDraft}
          hint={settings.autoDraft ? undefined : "Turn on Prepare drafts on its own first"}
          onChange={(event) => change({ maxDrafts: Number(event.target.value) })}
        >
          {DRAFT_COUNTS.map((count) => (
            <option key={count} value={count}>
              {count}
            </option>
          ))}
        </Select>
        <Toggle
          className="sm:col-span-2"
          label="Prepare drafts on its own"
          description={
            settings.autoDraft
              ? `For the best new postings of each run, up to ${plural(settings.maxDrafts, "draft", "drafts")}: a tailored resume and cover letter, waiting for your review.`
              : "Off. You choose which postings get a draft with Prepare draft."
          }
          checked={settings.autoDraft}
          onCheckedChange={(autoDraft) => change({ autoDraft })}
        />
      </Card>

      <p className="text-meta text-muted">
        One shared daily schedule for all enabled searches, in {timeZone}. Each search requests only the first page, sequentially. At most {AGENT_MAX_SOURCE_REQUESTS_PER_DAY} source requests in any rolling 24 hours, including failed attempts. A run waits until the whole batch fits. Across all searches, each run scores up to {plural(settings.maxScores, "posting", "postings")}
        {settings.autoDraft ? ` and prepares up to ${plural(settings.maxDrafts, "draft", "drafts")}` : ""}. Scoring and drafts go to your AI provider with your profile, as a new application does; the job source only gets the search.
      </p>
    </>
  );
}
