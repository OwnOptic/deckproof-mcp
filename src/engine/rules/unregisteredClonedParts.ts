/**
 * Rule 13 (error): parts that require an explicit <Override> content-type
 * (slides, layouts, masters, themes, notes, the presentation itself) must
 * have one. Rule 3 (contentTypesCompleteness) accepts a generic
 * <Default Extension="xml"> as coverage, but these part kinds each need a
 * SPECIFIC content type - a slide covered only by the xml Default is declared
 * as plain XML and won't be recognized as a slide. Cloning frequently copies
 * a layout/master part without adding its Override.
 */

import type { OpcPackage, Violation } from "../types.js";
import { listContentParts } from "../opcPackage.js";

export const RULE_ID = "unregisteredClonedParts";

const CT = "application/vnd.openxmlformats-officedocument";

/** part-path pattern -> the content type its <Override> must declare. */
const REQUIRED_OVERRIDES: Array<{ re: RegExp; contentType: string; label: string }> = [
  { re: /^ppt\/slides\/slide\d+\.xml$/, contentType: `${CT}.presentationml.slide+xml`, label: "slide" },
  { re: /^ppt\/slideLayouts\/slideLayout\d+\.xml$/, contentType: `${CT}.presentationml.slideLayout+xml`, label: "slide layout" },
  { re: /^ppt\/slideMasters\/slideMaster\d+\.xml$/, contentType: `${CT}.presentationml.slideMaster+xml`, label: "slide master" },
  { re: /^ppt\/notesSlides\/notesSlide\d+\.xml$/, contentType: `${CT}.presentationml.notesSlide+xml`, label: "notes slide" },
  { re: /^ppt\/notesMasters\/notesMaster\d+\.xml$/, contentType: `${CT}.presentationml.notesMaster+xml`, label: "notes master" },
  { re: /^ppt\/theme\/theme\d+\.xml$/, contentType: `${CT}.theme+xml`, label: "theme" },
];

export function checkUnregisteredClonedParts(pkg: OpcPackage): Violation[] {
  const overrides = new Map(
    pkg.contentTypes
      .filter((e) => e.kind === "override")
      .map((e) => [e.partName, e.contentType] as const)
  );

  const violations: Violation[] = [];
  for (const part of listContentParts(pkg)) {
    const spec = REQUIRED_OVERRIDES.find((s) => s.re.test(part));
    if (!spec) continue;
    const declared = overrides.get(`/${part}`);
    if (declared === undefined) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `${spec.label} part "${part}" has no <Override> in [Content_Types].xml; it needs one declaring "${spec.contentType}".`,
        partName: part,
      });
    } else if (declared !== spec.contentType) {
      violations.push({
        ruleId: RULE_ID, severity: "error",
        message: `${spec.label} part "${part}" is declared as "${declared}" but should be "${spec.contentType}".`,
        partName: part,
      });
    }
  }
  return violations;
}
