/**
 * Rule 10 (error): every slideMaster must reference exactly one theme, and
 * that theme part must exist. Cloning can leave a master pointing at a theme
 * that wasn't imported, breaking every color-mapped (clrMap) reference on
 * slides that inherit from it.
 */

import type { OpcPackage, Violation } from "../types.js";
import { partsWithPrefix, relationshipsOfType, resolveRelationshipTarget } from "../opcPackage.js";

export const RULE_ID = "themeColorMapBreakage";

export function checkThemeColorMapBreakage(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];
  for (const master of partsWithPrefix(pkg, "ppt/slideMasters/slideMaster")) {
    const themeRels = relationshipsOfType(pkg, master, "theme");
    if (themeRels.length === 0) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `Slide master ${master} references no theme; its color map (clrMap) cannot resolve.`,
        partName: master,
      });
      continue;
    }
    for (const rel of themeRels) {
      const themePart = resolveRelationshipTarget(master, rel.target);
      if (!pkg.parts.has(themePart)) {
        violations.push({
          ruleId: RULE_ID, severity: "error",
          message: `Slide master ${master} references theme "${rel.target}" which is missing from the package.`,
          partName: master,
        });
      }
    }
  }
  return violations;
}
