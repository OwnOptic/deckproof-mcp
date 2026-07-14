/**
 * Advisory audit: validation plus quality/portability metrics. Everything
 * here is derived from the same OpcPackage the rules run on - no new parsing
 * concepts, just aggregation for a human-facing report.
 */

import type { Violation } from "./types.js";
import { loadPackage, partsWithPrefix } from "./opcPackage.js";
import { runAllRules } from "./rules/index.js";
import { pictureAltStats } from "./rules/altTextCoverage.js";
import { checkOrphanedMedia } from "./rules/orphanedMedia.js";
import { checkBloatedMediaParts } from "./rules/bloatedMediaParts.js";

export type RiskLevel = "low" | "medium" | "high";

export interface AuditReport {
  valid: boolean;
  violations: Violation[];
  metrics: {
    slideCount: number;
    mediaCount: number;
    byteLength: number;
    altTextCoveragePct: number;
    pictureCount: number;
    orphanedMedia: string[];
    duplicateMediaGroups: number;
    portabilityRisk: {
      powerpoint: RiskLevel;
      keynote: RiskLevel;
      libreoffice: RiskLevel;
      googleSlides: RiskLevel;
    };
  };
}

export async function auditBuffer(bytes: Uint8Array | Buffer): Promise<AuditReport> {
  const pkg = await loadPackage(bytes);
  const violations = runAllRules(pkg);
  const errors = violations.filter((v) => v.severity === "error");
  const firedRuleIds = new Set(violations.map((v) => v.ruleId));

  const alt = pictureAltStats(pkg);
  const slideCount = partsWithPrefix(pkg, "ppt/slides/slide").filter((p) => /slide\d+\.xml$/.test(p)).length;
  const mediaCount = partsWithPrefix(pkg, "ppt/media/").length;
  const byteLength = bytes.byteLength;

  return {
    valid: errors.length === 0,
    violations,
    metrics: {
      slideCount,
      mediaCount,
      byteLength,
      pictureCount: alt.total,
      altTextCoveragePct: alt.total === 0 ? 100 : Math.round((alt.withAlt / alt.total) * 100),
      orphanedMedia: checkOrphanedMedia(pkg).map((v) => v.partName!).filter(Boolean),
      duplicateMediaGroups: checkBloatedMediaParts(pkg).length,
      portabilityRisk: assessPortability(errors.length > 0, firedRuleIds),
    },
  };
}

/**
 * Heuristic portability risk. Strict readers (Keynote, Google Slides import)
 * reject the structural bugs lenient readers (PowerPoint, LibreOffice)
 * silently self-heal - so the same file can be "fine" in one and broken in
 * another. Rule 2 (ID-lists), 4 (stale sldSz) and 6 (platform parts) are the
 * classic strict-reader tripwires.
 */
function assessPortability(hasErrors: boolean, fired: Set<string>): AuditReport["metrics"]["portabilityRisk"] {
  const strictTripwire =
    fired.has("presentationIdLists") ||
    fired.has("staleAttributes") ||
    fired.has("platformSpecificParts") ||
    fired.has("contentTypesCompleteness") ||
    fired.has("unregisteredClonedParts");

  const keynote: RiskLevel = strictTripwire ? "high" : hasErrors ? "medium" : "low";
  const googleSlides: RiskLevel = strictTripwire ? "high" : hasErrors ? "medium" : "low";
  const libreoffice: RiskLevel = hasErrors ? "medium" : "low";
  const powerpoint: RiskLevel = hasErrors ? "medium" : "low";
  return { powerpoint, keynote, libreoffice, googleSlides };
}
