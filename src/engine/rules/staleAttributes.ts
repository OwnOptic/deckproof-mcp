/**
 * Rule 4 (error): a <p:sldSz> whose `type` attribute contradicts its cx/cy
 * dimensions. This is the second half of the anthropics/skills#1167 root
 * cause: a deck resized to 16:9 keeps `type="screen4x3"`, a self-contradictory
 * element that Apple Keynote rejects outright.
 */

import type { OpcPackage, Violation } from "../types.js";
import { getPartXml } from "../opcPackage.js";
import { findAll, attr } from "../xml.js";

export const RULE_ID = "staleAttributes";

const PRESENTATION_PART = "ppt/presentation.xml";

// Standard named slide sizes -> exact EMU dimensions (ECMA-376).
const NAMED_SIZES: Record<string, { cx: number; cy: number }> = {
  screen4x3: { cx: 9144000, cy: 6858000 },
  screen16x9: { cx: 9144000, cy: 5143500 },
  screen16x10: { cx: 9144000, cy: 5715000 },
  a4: { cx: 9906000, cy: 6858000 },
  letter: { cx: 9144000, cy: 6858000 },
};

export function checkStaleAttributes(pkg: OpcPackage): Violation[] {
  const doc = getPartXml(pkg, PRESENTATION_PART);
  if (!doc) return [];

  const violations: Violation[] = [];
  for (const { node } of findAll(doc, "p:sldSz")) {
    const type = attr(node, "type");
    if (!type) continue;
    const expected = NAMED_SIZES[type.toLowerCase()];
    if (!expected) continue; // unknown/custom type name: nothing to contradict
    const cx = Number(attr(node, "cx"));
    const cy = Number(attr(node, "cy"));
    if (!cx || !cy) continue;
    // Compare aspect ratio (tolerant of minor rounding) rather than exact EMU.
    const actualRatio = cx / cy;
    const expectedRatio = expected.cx / expected.cy;
    if (Math.abs(actualRatio - expectedRatio) > 0.02) {
      violations.push({
        ruleId: RULE_ID,
        severity: "error",
        message: `<p:sldSz> declares type="${type}" but its dimensions (${cx}x${cy}) do not match that named size. Strict readers (e.g. Apple Keynote) reject this contradiction; remove the stale type attribute or fix the dimensions.`,
        partName: PRESENTATION_PART,
      });
    }
  }
  return violations;
}
