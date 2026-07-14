/**
 * Rule 6 (warning): Windows-only binary parts (printer settings) that have no
 * cross-platform equivalent. PowerPoint on Windows embeds
 * ppt/printerSettings/printerSettings*.bin; other implementations (Keynote,
 * Google Slides import) can choke on it. Not a hard spec violation, so a
 * warning - repair can strip it.
 */

import type { OpcPackage, Violation } from "../types.js";
import { listContentParts } from "../opcPackage.js";

export const RULE_ID = "platformSpecificParts";

const PATTERNS = [/^ppt\/printerSettings\/printerSettings\d+\.bin$/i];

export function checkPlatformSpecificParts(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];
  for (const part of listContentParts(pkg)) {
    if (PATTERNS.some((re) => re.test(part))) {
      violations.push({
        ruleId: RULE_ID, severity: "warning",
        message: `Part "${part}" is a Windows-only binary with no cross-platform equivalent; it can cause import failures in Keynote/Google Slides. Safe to remove.`,
        partName: part,
      });
    }
  }
  return violations;
}
