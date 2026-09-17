/**
 * PipelineTiming.js
 *
 * Tiny shared timer for Phase 4 measurement.
 * Writes one DB_log row per step via logAction (Phase3Processing.js).
 * Does not change pipeline data. Safe to leave on.
 *
 * Filter DB_log column Action = TIMING after a normal menu run.
 */
function pipelineTimingNow() {
  return Date.now();
}

/**
 * Log how long a step took.
 * stepName — exact function / landmark name
 * startedAtMs — value from pipelineTimingNow() taken just before the step
 * extra — optional short note (row counts, etc.)
 */
function pipelineTimingLog(stepName, startedAtMs, extra) {
  const ms = Date.now() - Number(startedAtMs || 0);
  const sec = (ms / 1000).toFixed(1);
  const details =
    String(stepName || "unknown") +
    " " +
    sec +
    "s (" +
    ms +
    "ms)" +
    (extra ? " | " + extra : "");
  logAction("TIMING", details);
  return ms;
}