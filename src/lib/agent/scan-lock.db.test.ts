import { Client } from "pg";
import { expect, it } from "vitest";

import { testDatabaseUrl } from "@/test/db";

import { withAgentScanLock } from "./scan-lock";

it("releases scan ownership when a process connection disappears, without a stale lease", async () => {
  const oldProcess = new Client({ connectionString: testDatabaseUrl() });
  await oldProcess.connect();
  try {
    await oldProcess.query("SELECT pg_advisory_lock(621746003)");
    let dispatched = false;
    expect(await withAgentScanLock(async () => { dispatched = true; }, testDatabaseUrl())).toBeNull();
    expect(dispatched).toBe(false);
  } finally { await oldProcess.end(); }
  expect(await withAgentScanLock(async () => "recovered", testDatabaseUrl())).toBe("recovered");
});

it("also releases ownership after an unexpected evaluation failure", async () => {
  await expect(withAgentScanLock(async () => { throw new Error("interrupted"); }, testDatabaseUrl())).rejects.toThrow("interrupted");
  expect(await withAgentScanLock(async () => "next scan", testDatabaseUrl())).toBe("next scan");
});
