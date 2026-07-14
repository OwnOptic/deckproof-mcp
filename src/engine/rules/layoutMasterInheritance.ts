/**
 * Rule 5 (error): the slide -> slideLayout -> slideMaster inheritance chain
 * must be intact. Every slide must relate to a layout, and every layout must
 * relate to a master, with the target parts present. A broken chain (common
 * after cloning) leaves a slide with no resolvable layout/master and renders
 * blank or fails to open.
 */

import type { OpcPackage, Violation } from "../types.js";
import { partsWithPrefix, relationshipsOfType, resolveRelationshipTarget } from "../opcPackage.js";

export const RULE_ID = "layoutMasterInheritance";

export function checkLayoutMasterInheritance(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];

  for (const slide of partsWithPrefix(pkg, "ppt/slides/slide")) {
    const layoutRels = relationshipsOfType(pkg, slide, "slideLayout");
    if (layoutRels.length === 0) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `Slide ${slide} has no slideLayout relationship; every slide must reference a layout.`,
        partName: slide,
      });
      continue;
    }
    for (const rel of layoutRels) {
      const layoutPart = resolveRelationshipTarget(slide, rel.target);
      if (!pkg.parts.has(layoutPart)) continue; // dangling caught by rule 1
      const masterRels = relationshipsOfType(pkg, layoutPart, "slideMaster");
      if (masterRels.length === 0) {
        violations.push({
          ruleId: RULE_ID, severity: "error",
          message: `Layout ${layoutPart} (used by ${slide}) has no slideMaster relationship; the layout->master link is broken.`,
          partName: layoutPart,
        });
      }
    }
  }

  return violations;
}
