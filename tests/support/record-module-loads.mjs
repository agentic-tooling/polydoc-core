/**
 * Module customization hooks that record every module URL Node loads.
 *
 * Hooks run on their own thread, so the record is appended synchronously to a
 * file rather than posted over a message port: there is nothing to drain, and
 * the file is complete by the time the probe process reads it back.
 */
import { appendFileSync } from "node:fs";

let recordPath = "";

export async function initialize(data) {
  recordPath = data.recordPath;
}

export async function load(url, context, nextLoad) {
  appendFileSync(recordPath, `${url}\n`);
  return nextLoad(url, context);
}
