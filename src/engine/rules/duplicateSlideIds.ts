/**
 * Rule 9 (error): every <p:sldId> in presentation.xml's <p:sldIdLst> must
 * have a unique `id` and a unique `r:id`. Cloning can produce colliding slide
 * IDs, which makes PowerPoint drop or merge slides on open.
 */

import type { OpcPackage, Violation } from "../types.js";
import { getPartXml } from "../opcPackage.js";
import { findAll, attr } from "../xml.js";

export const RULE_ID = "duplicateSlideIds";

const PRESENTATION_PART = "ppt/presentation.xml";

export function checkDuplicateSlideIds(pkg: OpcPackage): Violation[] {
  const doc = getPartXml(pkg, PRESENTATION_PART);
  if (!doc) return [];

  const violations: Violation[] = [];
  const seenIds = new Set<string>();
  const seenRIds = new Set<string>();

  for (const { node } of findAll(doc, "p:sldId")) {
    const id = attr(node, "id");
    const rId = attr(node, "r:id");
    if (id !== undefined) {
      if (seenIds.has(id)) {
        violations.push({
          ruleId: RULE_ID, severity: "error",
          message: `Duplicate slide id="${id}" in <p:sldIdLst>; slide ids must be unique.`,
          partName: PRESENTATION_PART,
        });
      }
      seenIds.add(id);
    }
    if (rId !== undefined) {
      if (seenRIds.has(rId)) {
        violations.push({
          ruleId: RULE_ID, severity: "error",
          message: `Duplicate slide r:id="${rId}" in <p:sldIdLst>; each slide entry must reference a distinct relationship.`,
          partName: PRESENTATION_PART,
        });
      }
      seenRIds.add(rId);
    }
  }

  return violations;
}
