/**
 * Rule 12 (error): within any single .rels part, every <Relationship> Id must
 * be unique. Duplicate rIds (a common clone artifact) make target resolution
 * ambiguous and can silently mis-wire slides to the wrong parts.
 */

import type { OpcPackage, Violation } from "../types.js";
import { relsPathFor } from "../opcPackage.js";

export const RULE_ID = "staleRelationshipIds";

export function checkStaleRelationshipIds(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];
  for (const [sourcePart, rels] of pkg.relsByPart) {
    const seen = new Set<string>();
    for (const rel of rels) {
      if (seen.has(rel.id)) {
        violations.push({
          ruleId: RULE_ID, severity: "error",
          message: `Duplicate relationship Id "${rel.id}" in ${relsPathFor(sourcePart)}; relationship ids must be unique within a .rels part.`,
          partName: sourcePart,
        });
      }
      seen.add(rel.id);
    }
  }
  return violations;
}
