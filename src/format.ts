/**
 * Human-readable text summaries. Every tool returns a compact text block
 * alongside its structured content so results read well in any client.
 */

import type { Violation } from "./schemas.js";

export function formatViolations(violations: Violation[]): string {
  if (violations.length === 0) return "No violations found.";
  const errors = violations.filter((v) => v.severity === "error");
  const warnings = violations.filter((v) => v.severity === "warning");
  const lines: string[] = [];
  if (errors.length > 0) {
    lines.push(`${errors.length} error(s):`);
    for (const v of errors) {
      lines.push(`  [${v.ruleId}] ${v.message}${v.partName ? ` (${v.partName})` : ""}`);
    }
  }
  if (warnings.length > 0) {
    lines.push(`${warnings.length} warning(s):`);
    for (const v of warnings) {
      lines.push(`  [${v.ruleId}] ${v.message}${v.partName ? ` (${v.partName})` : ""}`);
    }
  }
  return lines.join("\n");
}
