/**
 * Rule 8 (warning): pictures without alt text. Each <p:pic> should carry a
 * `descr` on its <p:cNvPr> for accessibility. Non-blocking, but reported here
 * and surfaced as a coverage percentage by pptx_audit.
 */

import type { OpcPackage, Violation } from "../types.js";
import { partsWithPrefix, getPartXml } from "../opcPackage.js";
import { findAll, attr } from "../xml.js";

export const RULE_ID = "altTextCoverage";

/** Count pictures and how many have non-empty alt text, across all slides. */
export function pictureAltStats(pkg: OpcPackage): { total: number; withAlt: number } {
  let total = 0;
  let withAlt = 0;
  for (const slide of partsWithPrefix(pkg, "ppt/slides/slide")) {
    const doc = getPartXml(pkg, slide);
    if (!doc) continue;
    for (const pic of findAll(doc, "p:pic")) {
      const cNvPr = findAll(pic.children, "p:cNvPr")[0];
      if (!cNvPr) continue;
      total += 1;
      const descr = attr(cNvPr.node, "descr");
      if (descr && descr.trim().length > 0) withAlt += 1;
    }
  }
  return { total, withAlt };
}

export function checkAltTextCoverage(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];
  for (const slide of partsWithPrefix(pkg, "ppt/slides/slide")) {
    const doc = getPartXml(pkg, slide);
    if (!doc) continue;
    for (const pic of findAll(doc, "p:pic")) {
      const cNvPr = findAll(pic.children, "p:cNvPr")[0];
      if (!cNvPr) continue;
      const descr = attr(cNvPr.node, "descr");
      if (!descr || descr.trim().length === 0) {
        const name = attr(cNvPr.node, "name") ?? "(unnamed)";
        violations.push({
          ruleId: RULE_ID, severity: "warning",
          message: `Picture "${name}" in ${slide} has no alt text (descr). Add one for accessibility.`,
          partName: slide,
        });
      }
    }
  }
  return violations;
}
