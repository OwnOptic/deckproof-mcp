/**
 * Rule registry. Each rule is a pure `(pkg) => Violation[]` function; adding
 * a new rule means adding one file here and one line below - nothing else
 * in validate.ts, sanitize.ts, or the tools needs to change.
 *
 * Rules 1-6 and 9-13 are error-severity (they gate `valid`); rules 7, 8, 14
 * are advisory warnings surfaced richly by pptx_audit.
 */

import type { OpcPackage, Violation } from "../types.js";
import { checkDanglingRelationships } from "./danglingRelationships.js";
import { checkPresentationIdLists } from "./presentationIdLists.js";
import { checkContentTypesCompleteness } from "./contentTypesCompleteness.js";
import { checkStaleAttributes } from "./staleAttributes.js";
import { checkLayoutMasterInheritance } from "./layoutMasterInheritance.js";
import { checkPlatformSpecificParts } from "./platformSpecificParts.js";
import { checkOrphanedMedia } from "./orphanedMedia.js";
import { checkAltTextCoverage } from "./altTextCoverage.js";
import { checkDuplicateSlideIds } from "./duplicateSlideIds.js";
import { checkThemeColorMapBreakage } from "./themeColorMapBreakage.js";
import { checkOrphanedMasterLayoutChains } from "./orphanedMasterLayoutChains.js";
import { checkStaleRelationshipIds } from "./staleRelationshipIds.js";
import { checkUnregisteredClonedParts } from "./unregisteredClonedParts.js";
import { checkBloatedMediaParts } from "./bloatedMediaParts.js";

export type Rule = (pkg: OpcPackage) => Violation[];

export const rules: Rule[] = [
  checkDanglingRelationships, // 1
  checkPresentationIdLists, // 2
  checkContentTypesCompleteness, // 3
  checkStaleAttributes, // 4
  checkLayoutMasterInheritance, // 5
  checkPlatformSpecificParts, // 6
  checkOrphanedMedia, // 7
  checkAltTextCoverage, // 8
  checkDuplicateSlideIds, // 9
  checkThemeColorMapBreakage, // 10
  checkOrphanedMasterLayoutChains, // 11
  checkStaleRelationshipIds, // 12
  checkUnregisteredClonedParts, // 13
  checkBloatedMediaParts, // 14
];

export function runAllRules(pkg: OpcPackage): Violation[] {
  return rules.flatMap((rule) => rule(pkg));
}
