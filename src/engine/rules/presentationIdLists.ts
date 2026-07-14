/**
 * Rule 2 (error): every slideMaster/slide/notesMaster relationship declared
 * on ppt/presentation.xml must be registered in the corresponding ID-list
 * element (<p:sldMasterIdLst>, <p:sldIdLst>, <p:notesMasterIdLst> -
 * ECMA-376 sec 19.2.1.30). A relationship + part existing but NOT listed is
 * exactly the bug behind anthropics/skills#1167: PowerPoint and LibreOffice
 * tolerate the omission and silently self-heal on save, but Apple Keynote
 * (strict) rejects the file outright.
 */

import type { OpcPackage, Violation } from "../types.js";
import { getPartXml } from "../opcPackage.js";
import { findAll, attr } from "../xml.js";

export const RULE_ID = "presentationIdLists";

const PRESENTATION_PART = "ppt/presentation.xml";

interface IdListSpec {
  relTypeSuffix: string;
  containerTag: string;
  itemTag: string;
  label: string;
}

const ID_LISTS: IdListSpec[] = [
  { relTypeSuffix: "/slideMaster", containerTag: "p:sldMasterIdLst", itemTag: "p:sldMasterId", label: "slide master" },
  { relTypeSuffix: "/slide", containerTag: "p:sldIdLst", itemTag: "p:sldId", label: "slide" },
  { relTypeSuffix: "/notesMaster", containerTag: "p:notesMasterIdLst", itemTag: "p:notesMasterId", label: "notes master" },
];

export function checkPresentationIdLists(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];
  const doc = getPartXml(pkg, PRESENTATION_PART);
  if (!doc) return violations; // no presentation.xml at all is caught by other checks

  const presentationRels = pkg.relsByPart.get(PRESENTATION_PART) ?? [];

  for (const spec of ID_LISTS) {
    // Relationships of this type that should appear in the list. A "/slide"
    // suffix match must exclude "/slideMaster" and "/slideLayout", which also
    // end in a superstring - use an exact trailing-segment comparison.
    const expectedIds = presentationRels
      .filter((r) => r.type.split("/").pop() === spec.relTypeSuffix.slice(1))
      .map((r) => r.id);
    if (expectedIds.length === 0) continue;

    const listedIds = new Set(
      findAll(doc, spec.itemTag).map(({ node }) => attr(node, "r:id")).filter(Boolean)
    );

    for (const id of expectedIds) {
      if (!listedIds.has(id)) {
        violations.push({
          ruleId: RULE_ID,
          severity: "error",
          message: `${PRESENTATION_PART} has a ${spec.label} relationship (${id}) that is not registered in <${spec.containerTag}>. PowerPoint/LibreOffice may self-heal this silently; strict readers (e.g. Apple Keynote) will reject the file.`,
          partName: PRESENTATION_PART,
        });
      }
    }
  }

  return violations;
}
