/**
 * Rule 14 (warning): duplicate media parts (identical bytes stored under
 * multiple names). Repeated clone operations copy the same image many times,
 * bloating the file. Non-blocking - repair can deduplicate.
 */

import { createHash } from "node:crypto";
import type { OpcPackage, Violation } from "../types.js";
import { partsWithPrefix } from "../opcPackage.js";

export const RULE_ID = "bloatedMediaParts";

export function checkBloatedMediaParts(pkg: OpcPackage): Violation[] {
  const byHash = new Map<string, string[]>();
  for (const media of partsWithPrefix(pkg, "ppt/media/")) {
    const bytes = pkg.parts.get(media);
    if (!bytes) continue;
    const hash = createHash("md5").update(bytes).digest("hex");
    const group = byHash.get(hash) ?? [];
    group.push(media);
    byHash.set(hash, group);
  }

  const violations: Violation[] = [];
  for (const group of byHash.values()) {
    if (group.length > 1) {
      violations.push({
        ruleId: RULE_ID, severity: "warning",
        message: `Duplicate media (identical bytes) stored ${group.length} times: ${group.join(", ")}. Deduplicate to reduce file size.`,
        partName: group[0],
      });
    }
  }
  return violations;
}
