/**
 * Rule 3 (error): every part in the package must have a content type -
 * either a <Default> entry matching its file extension, or an <Override>
 * entry matching its exact part name. A part with neither is invisible to
 * strict OOXML consumers (e.g. Apache POI throws InvalidFormatException);
 * lenient consumers (PowerPoint) often guess from the extension and open it
 * anyway, hiding the defect.
 *
 * This checks *presence* only, not whether an existing declared content type
 * is the semantically "correct" one for that part kind - verifying the
 * latter would need a canonical per-part-kind content-type map, which is a
 * larger undertaking deferred past v1.
 */

import type { OpcPackage, Violation } from "../types.js";
import { listContentParts } from "../opcPackage.js";

export const RULE_ID = "contentTypesCompleteness";

function extensionOf(partName: string): string {
  const dot = partName.lastIndexOf(".");
  return dot === -1 ? "" : partName.slice(dot + 1).toLowerCase();
}

export function checkContentTypesCompleteness(pkg: OpcPackage): Violation[] {
  const violations: Violation[] = [];

  const defaultExtensions = new Set(
    pkg.contentTypes
      .filter((e) => e.kind === "default")
      .map((e) => e.extension.toLowerCase())
  );
  const overridePartNames = new Set(
    pkg.contentTypes.filter((e) => e.kind === "override").map((e) => e.partName)
  );

  for (const partName of listContentParts(pkg)) {
    const ext = extensionOf(partName);
    const hasDefault = ext !== "" && defaultExtensions.has(ext);
    const hasOverride = overridePartNames.has(`/${partName}`);
    if (!hasDefault && !hasOverride) {
      violations.push({
        ruleId: RULE_ID,
        severity: "error",
        message: `Part "${partName}" has no content-type declaration - no <Default Extension="${ext}"> and no <Override PartName="/${partName}"> in [Content_Types].xml.`,
        partName,
      });
    }
  }

  return violations;
}
