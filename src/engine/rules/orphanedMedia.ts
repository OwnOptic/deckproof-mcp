/**
 * Rule 7 (warning): media parts (ppt/media/*) not referenced by any
 * relationship. Orphaned media inflates file size and signals a broken
 * import; not a hard spec violation, so a warning - repair can prune it.
 */

import type { OpcPackage, Violation } from "../types.js";
import { partsWithPrefix, resolvedInternalTargets } from "../opcPackage.js";

export const RULE_ID = "orphanedMedia";

export function checkOrphanedMedia(pkg: OpcPackage): Violation[] {
  const referenced = resolvedInternalTargets(pkg);
  const violations: Violation[] = [];
  for (const media of partsWithPrefix(pkg, "ppt/media/")) {
    if (!referenced.has(media)) {
      violations.push({
        ruleId: RULE_ID, severity: "warning",
        message: `Media part "${media}" is not referenced by any relationship (orphaned; inflates file size).`,
        partName: media,
      });
    }
  }
  return violations;
}
