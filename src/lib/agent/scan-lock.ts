import "server-only";

import { Client } from "pg";

import { env } from "@/lib/env";
import { logServerError } from "@/lib/log";
import { prisma } from "@/lib/prisma";

export const AGENT_SCAN_LOCK_ID = 621746003;

/** Read-only screen status. Never acquire the execution lock just to render:
 * a Server Action refresh can race the scan job it has just queued. */
export async function isAgentScanRunning(): Promise<boolean> {
  const [row] = await prisma.$queryRaw<{ running: boolean }[]>`SELECT EXISTS (
    SELECT 1 FROM pg_locks WHERE locktype = 'advisory' AND classid = 0
    AND objid = ${AGENT_SCAN_LOCK_ID}::oid AND objsubid = 1 AND granted
    AND database = (SELECT oid FROM pg_database WHERE datname = current_database())
  ) AS running`;
  return row.running;
}

/** A dedicated session lock, not an idle Prisma transaction with a timeout.
 * It spans provider calls and evaluation, and dies with the connection on a
 * process restart. A slow model must not outlive a transaction timeout and
 * silently lose the scan lock. The optional URL also isolates database tests.
 * This serializes scans only, not imports/wipes or independent draft jobs.
 */
export async function withAgentScanLock<T>(work: (signal: AbortSignal) => Promise<T>, connectionString = env.DATABASE_URL): Promise<T | null> {
  const client = new Client({ connectionString });
  const controller = new AbortController();
  client.on("error", () => controller.abort());
  try {
    await client.connect();
    const result = await client.query<{ locked: boolean }>("SELECT pg_try_advisory_lock($1::bigint) AS locked", [AGENT_SCAN_LOCK_ID]);
    if (!result.rows[0].locked) return null;
    // Closing this dedicated connection in finally releases the session lock.
    return await work(controller.signal);
  } finally {
    // Also clean up a failed connect, and never mask the useful work/connect
    // error with a secondary connection-close failure.
    await client.end().catch((error: unknown) => logServerError("job agent scan lock cleanup", error));
  }
}
