/**
 * Shared engine types. An OpcPackage is the single abstraction every rule,
 * the sanitizer, and the creators operate on - no rule ever touches raw
 * zip bytes or XML strings directly.
 */

export interface Violation {
  ruleId: string;
  severity: "error" | "warning";
  message: string;
  partName?: string;
}

export interface Change {
  ruleId: string;
  description: string;
  partName?: string;
}

/** One <Relationship> entry from a .rels part. */
export interface OpcRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: "Internal" | "External";
}

/** One entry from [Content_Types].xml - either a Default (by extension) or an Override (by exact part name). */
export interface ContentTypeDefault {
  kind: "default";
  extension: string;
  contentType: string;
}
export interface ContentTypeOverride {
  kind: "override";
  partName: string;
  contentType: string;
}
export type ContentTypeEntry = ContentTypeDefault | ContentTypeOverride;

/**
 * In-memory model of an Open Packaging Conventions (.pptx) package.
 * `parts` holds every part's raw bytes keyed by its zip-internal path
 * (e.g. "ppt/presentation.xml", "ppt/media/image1.png").
 * `relsByPart` and `contentTypes` are parsed views kept in sync with the
 * corresponding `.rels` / `[Content_Types].xml` parts in `parts` - callers
 * mutate via the helper functions in opcPackage.ts, never the maps directly,
 * so the two stay consistent.
 */
export interface OpcPackage {
  parts: Map<string, Uint8Array>;
  /** Relationships keyed by the .rels part's *source* (e.g. "ppt/presentation.xml" -> its relationships). */
  relsByPart: Map<string, OpcRelationship[]>;
  contentTypes: ContentTypeEntry[];
}

export type LayoutArchetype =
  | "cover"
  | "agenda"
  | "contentBullets"
  | "twoColumns"
  | "comparisonTable"
  | "dataTable"
  | "timeline"
  | "statsBanner"
  | "quote"
  | "cardGrid"
  | "sectionDivider"
  | "orgChart"
  | "matrixQuadrant"
  | "closingNextSteps"
  | "verticalSteps";
