/**
 * Rule 1 (error): every internal relationship's Target must resolve to a
 * part that actually exists in the package. A relationship pointing at a
 * missing part is the single most common way a .pptx becomes unopenable -
 * whatever referenced it (a slide, a master, a layout) will fail to load.
 *
 * External relationships (hyperlinks, etc.) are exempt - their target is a
 * URL, not a package part.
 */

import type { OpcPackage, Violation } from "../types.js";
import { resolveRelationshipTarget } from "../opcPackage.js";

export const RULE_ID = "danglingRelationships";

export function checkDanglingRelationships(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];

  for (const [sourcePart, rels] of pkg.relsByPart) {
    for (const rel of rels) {
      if (rel.targetMode === "External") continue;
      const resolved = resolveRelationshipTarget(sourcePart, rel.target);
      if (!pkg.parts.has(resolved)) {
        violations.push({
          ruleId: RULE_ID,
          severity: "error",
          message: `Relationship ${rel.id} in ${sourcePart || "(root)"} targets "${rel.target}" (resolved to "${resolved}"), which does not exist in the package.`,
          partName: sourcePart,
        });
      }
    }
  }

  return violations;
}
