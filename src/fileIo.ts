/**
 * Reads a `FileSource` (local path or base64) into bytes, with the security
 * guards that matter for a server ingesting untrusted uploads:
 *  - `path` (reads a local file) is honored ONLY on the local stdio transport;
 *    over HTTP it is rejected, since an attacker-controlled path is a
 *    local-file-read / SSRF vector.
 *  - inputs over MAX_INPUT_BYTES are rejected before any parsing.
 * Filesystem access is a tool-layer concern, deliberately kept out of engine/.
 */

import { readFile } from "node:fs/promises";
import type { FileSource } from "./schemas.js";

export const MAX_INPUT_BYTES = 50 * 1024 * 1024; // 50 MB

export interface ReadOptions {
  /** True only under the local stdio transport (set by the server bootstrap). */
  allowLocalPaths: boolean;
}

function checkSize(buf: Buffer): Buffer {
  if (buf.byteLength > MAX_INPUT_BYTES) {
    throw new Error(
      `Input is ${(buf.byteLength / (1024 * 1024)).toFixed(1)} MB, over the ${Math.round(MAX_INPUT_BYTES / (1024 * 1024))} MB limit.`
    );
  }
  return buf;
}

export async function readFileSource(source: FileSource, opts: ReadOptions): Promise<Buffer> {
  if (source.path) {
    if (!opts.allowLocalPaths) {
      throw new Error(
        "The `path` input is only available on the local stdio transport. Over HTTP, pass the file as base64 instead."
      );
    }
    return checkSize(await readFile(source.path));
  }
  if (source.base64) {
    return checkSize(Buffer.from(source.base64, "base64"));
  }
  throw new Error("FileSource must provide either `path` or `base64`.");
}
