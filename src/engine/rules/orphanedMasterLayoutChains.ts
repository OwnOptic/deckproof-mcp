/**
 * Rule 11 (error): slide masters and layouts must be properly wired.
 * - Every slideMaster must be registered in presentation.xml's
 *   <p:sldMasterIdLst> (an unregistered master is dead weight and can break
 *   layout resolution).
 * - Every slideLayout must relate back to a slideMaster (an orphaned layout,
 *   imported without its parent master during a clone, cannot resolve).
 */

import type { OpcPackage, Violation } from "../types.js";
import {
  partsWithPrefix,
  relationshipsOfType,
  resolveRelationshipTarget,
  resolvedInternalTargets,
  getPartXml,
} from "../opcPackage.js";

export const RULE_ID = "orphanedMasterLayoutChains";

const PRESENTATION_PART = "ppt/presentation.xml";

export function checkOrphanedMasterLayoutChains(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];

  // Masters registered via presentation.xml -> sldMasterIdLst relationships.
  const registeredMasters = new Set<string>();
  if (getPartXml(pkg, PRESENTATION_PART)) {
    for (const rel of relationshipsOfType(pkg, PRESENTATION_PART, "slideMaster")) {
      registeredMasters.add(resolveRelationshipTarget(PRESENTATION_PART, rel.target));
    }
  }
  for (const master of partsWithPrefix(pkg, "ppt/slideMasters/slideMaster")) {
    if (!registeredMasters.has(master)) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `Slide master ${master} is not registered in presentation.xml <p:sldMasterIdLst>.`,
        partName: master,
      });
    }
  }

  // Layouts must be linked from some master (they appear as a resolved target
  // of a slideMaster relationship) AND relate back to a master.
  const referenced = resolvedInternalTargets(pkg);
  for (const layout of partsWithPrefix(pkg, "ppt/slideLayouts/slideLayout")) {
    const masterRels = relationshipsOfType(pkg, layout, "slideMaster");
    if (masterRels.length === 0) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `Slide layout ${layout} has no relationship to a slide master (orphaned layout).`,
        partName: layout,
      });
    }
    if (!referenced.has(layout)) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `Slide layout ${layout} is not referenced by any master or slide (orphaned).`,
        partName: layout,
      });
    }
  }

  return violations;
}
