/**
 * Structural validation: load a .pptx into an OpcPackage, run every rule,
 * and report whether it's `valid` (no error-severity violations) alongside
 * the full violation list (errors and warnings both).
 */

import type { Violation } from "./types.js";
import { loadPackage } from "./opcPackage.js";
import { runAllRules } from "./rules/index.js";

export interface ValidateResult {
  valid: boolean;
  violations: Violation[];
}

export async function validateBuffer(bytes: Uint8Array | Buffer): Promise<ValidateResult> {
  const pkg = await loadPackage(bytes);
  const violations = runAllRules(pkg);
  const valid = !violations.some((v) => v.severity === "error");
  return { valid, violations };
}
