/**
 * Repair pass: a from-scratch reimplementation of generic OOXML/OPC
 * sanitation (no code from any private tooling). Each fixer is a pure-ish
 * `(pkg) => Change[]` that mutates the package in place and reports what it
 * changed. `sanitizeBuffer` composes them, re-serializes, and re-validates so
 * the caller sees both what was fixed and what could not be fixed safely.
 *
 * Fixers only make SAFE changes: they remove things that point nowhere,
 * register things that exist, and normalize clearly-invalid values. They
 * never invent content or guess at a user's intent.
 */

import type { OpcPackage, Change, Violation } from "./types.js";
import {
  loadPackage,
  savePackage,
  listContentParts,
  partsWithPrefix,
  resolvedInternalTargets,
  resolveRelationshipTarget,
  deletePart,
  getPartXml,
  setPartXml,
  isRelsPart,
} from "./opcPackage.js";
import { findAll, attr, type XmlDocument, type XmlNode } from "./xml.js";
import { runAllRules } from "./rules/index.js";
import { createHash } from "node:crypto";

const CT = "application/vnd.openxmlformats-officedocument";
const PRESENTATION_PART = "ppt/presentation.xml";
const CONTENT_TYPES_PART = "[Content_Types].xml";

const REQUIRED_OVERRIDES: Array<{ re: RegExp; contentType: string }> = [
  { re: /^ppt\/slides\/slide\d+\.xml$/, contentType: `${CT}.presentationml.slide+xml` },
  { re: /^ppt\/slideLayouts\/slideLayout\d+\.xml$/, contentType: `${CT}.presentationml.slideLayout+xml` },
  { re: /^ppt\/slideMasters\/slideMaster\d+\.xml$/, contentType: `${CT}.presentationml.slideMaster+xml` },
  { re: /^ppt\/notesSlides\/notesSlide\d+\.xml$/, contentType: `${CT}.presentationml.notesSlide+xml` },
  { re: /^ppt\/notesMasters\/notesMaster\d+\.xml$/, contentType: `${CT}.presentationml.notesMaster+xml` },
  { re: /^ppt\/theme\/theme\d+\.xml$/, contentType: `${CT}.theme+xml` },
];

// --- individual fixers ------------------------------------------------------

/** Add missing/incorrect <Override> content types for slides, layouts, masters, themes, notes. */
function fixMissingOverrides(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  const overrideIndex = new Map(
    pkg.contentTypes
      .map((e, i) => [e, i] as const)
      .filter(([e]) => e.kind === "override")
      .map(([e, i]) => [(e as { partName: string }).partName, i] as const)
  );
  for (const part of listContentParts(pkg)) {
    const spec = REQUIRED_OVERRIDES.find((s) => s.re.test(part));
    if (!spec) continue;
    const partName = `/${part}`;
    const existingIdx = overrideIndex.get(partName);
    if (existingIdx === undefined) {
      pkg.contentTypes.push({ kind: "override", partName, contentType: spec.contentType });
      changes.push({ ruleId: "unregisteredClonedParts", description: `Added Content-Types Override for ${part}`, partName: part });
    } else {
      const entry = pkg.contentTypes[existingIdx];
      if (entry.kind === "override" && entry.contentType !== spec.contentType) {
        entry.contentType = spec.contentType;
        changes.push({ ruleId: "unregisteredClonedParts", description: `Corrected content type for ${part}`, partName: part });
      }
    }
  }
  return changes;
}

/** Remove <Override> entries whose target part no longer exists. */
function fixDanglingOverrides(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  const before = pkg.contentTypes.length;
  pkg.contentTypes = pkg.contentTypes.filter((e) => {
    if (e.kind !== "override") return true;
    const part = e.partName.replace(/^\//, "");
    if (pkg.parts.has(part)) return true;
    changes.push({ ruleId: "contentTypesCompleteness", description: `Removed dangling Content-Types Override for ${e.partName}` });
    return false;
  });
  void before;
  return changes;
}

/** Delete Windows-only printer-settings binaries plus any relationships to them. */
function fixPlatformSpecificParts(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  for (const part of listContentParts(pkg)) {
    if (/^ppt\/printerSettings\/printerSettings\d+\.bin$/i.test(part)) {
      deletePart(pkg, part);
      changes.push({ ruleId: "platformSpecificParts", description: `Removed platform-specific part ${part}`, partName: part });
    }
  }
  // Drop any relationships that now point at a removed printerSettings part.
  for (const [source, rels] of pkg.relsByPart) {
    const kept = rels.filter((r) => {
      if (r.targetMode === "External") return true;
      const resolved = resolveRelationshipTarget(source, r.target);
      return !/printerSettings\d+\.bin$/i.test(resolved) || pkg.parts.has(resolved);
    });
    if (kept.length !== rels.length) pkg.relsByPart.set(source, kept);
  }
  return changes;
}

/** Delete media parts nothing references. */
function fixOrphanedMedia(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  const referenced = resolvedInternalTargets(pkg);
  for (const media of partsWithPrefix(pkg, "ppt/media/")) {
    if (!referenced.has(media)) {
      deletePart(pkg, media);
      changes.push({ ruleId: "orphanedMedia", description: `Removed orphaned media ${media}`, partName: media });
    }
  }
  return changes;
}

/** Deduplicate identical media: keep one copy, repoint relationships, delete the rest. */
function fixBloatedMedia(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  const byHash = new Map<string, string[]>();
  for (const media of partsWithPrefix(pkg, "ppt/media/")) {
    const bytes = pkg.parts.get(media);
    if (!bytes) continue;
    const hash = createHash("md5").update(bytes).digest("hex");
    (byHash.get(hash) ?? byHash.set(hash, []).get(hash)!).push(media);
  }
  for (const group of byHash.values()) {
    if (group.length < 2) continue;
    const [keep, ...dups] = group.sort();
    const keepFile = keep.split("/").pop()!;
    for (const dup of dups) {
      const dupFile = dup.split("/").pop()!;
      for (const [, rels] of pkg.relsByPart) {
        for (const rel of rels) {
          if (rel.target.split("/").pop() === dupFile) {
            rel.target = rel.target.replace(dupFile, keepFile);
          }
        }
      }
      deletePart(pkg, dup);
      changes.push({ ruleId: "bloatedMediaParts", description: `Deduplicated media ${dup} -> ${keep}`, partName: dup });
    }
  }
  return changes;
}

/** Remove relationships whose internal target is missing, plus any presentation ID-list entry that referenced them. */
function fixDanglingRelationships(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  const removedRIds = new Set<string>();
  for (const [source, rels] of pkg.relsByPart) {
    const kept = rels.filter((r) => {
      if (r.targetMode === "External") return true;
      const resolved = resolveRelationshipTarget(source, r.target);
      if (pkg.parts.has(resolved)) return true;
      if (source === PRESENTATION_PART) removedRIds.add(r.id);
      changes.push({ ruleId: "danglingRelationships", description: `Removed dangling relationship ${r.id} (-> ${r.target}) from ${source || "(root)"}`, partName: source });
      return false;
    });
    if (kept.length !== rels.length) pkg.relsByPart.set(source, kept);
  }
  if (removedRIds.size > 0) removeIdListEntries(pkg, removedRIds, changes);
  return changes;
}

/** Backfill presentation ID-list entries for slide/master/notes relationships that exist but aren't listed. */
function fixPresentationIdLists(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  const doc = getPartXml(pkg, PRESENTATION_PART);
  if (!doc) return changes;
  const rels = pkg.relsByPart.get(PRESENTATION_PART) ?? [];

  const specs = [
    { suffix: "slideMaster", listTag: "p:sldMasterIdLst", itemTag: "p:sldMasterId" },
    { suffix: "notesMaster", listTag: "p:notesMasterIdLst", itemTag: "p:notesMasterId" },
    { suffix: "slide", listTag: "p:sldIdLst", itemTag: "p:sldId" },
  ];

  let nextId = maxIdInPresentation(doc) + 1;
  let mutated = false;

  for (const spec of specs) {
    const relIds = rels.filter((r) => r.type.split("/").pop() === spec.suffix).map((r) => r.id);
    if (relIds.length === 0) continue;
    const listChildren = ensureIdList(doc, spec.listTag);
    const listed = new Set(
      findAll(listChildren, spec.itemTag).map(({ node }) => attr(node, "r:id")).filter(Boolean)
    );
    for (const rId of relIds) {
      if (listed.has(rId)) continue;
      const entry: XmlNode = { [spec.itemTag]: [], ":@": { "@_id": String(nextId), "@_r:id": rId } } as XmlNode;
      // notesMasterId has no numeric id attribute in the schema; only sldId/sldMasterId do.
      if (spec.itemTag === "p:notesMasterId") delete (entry[":@"] as Record<string, string>)["@_id"];
      else nextId += 1;
      listChildren.push(entry);
      mutated = true;
      changes.push({ ruleId: "presentationIdLists", description: `Registered ${spec.suffix} relationship ${rId} in <${spec.listTag}>`, partName: PRESENTATION_PART });
    }
  }

  if (mutated) setPartXml(pkg, PRESENTATION_PART, doc);
  return changes;
}

/** Clamp negative DrawingML extents (<a:ext cx cy>) to zero across slide/layout/master parts. */
function fixNegativeExtents(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  for (const part of listContentParts(pkg)) {
    if (!/^ppt\/(slides|slideLayouts|slideMasters)\/.*\.xml$/.test(part)) continue;
    const doc = getPartXml(pkg, part);
    if (!doc) continue;
    let mutated = false;
    for (const { node } of findAll(doc, "a:ext")) {
      const bag = (node as Record<string, unknown>)[":@"] as Record<string, string> | undefined;
      if (!bag) continue;
      for (const dim of ["@_cx", "@_cy"]) {
        if (bag[dim] !== undefined && Number(bag[dim]) < 0) {
          bag[dim] = "0";
          mutated = true;
        }
      }
    }
    if (mutated) {
      setPartXml(pkg, part, doc);
      changes.push({ ruleId: "staleAttributes", description: `Clamped negative DrawingML extent(s) in ${part}`, partName: part });
    }
  }
  return changes;
}

// --- XML mutation helpers ---------------------------------------------------

function maxIdInPresentation(doc: XmlDocument): number {
  let max = 255; // slide ids conventionally start at 256
  for (const tag of ["p:sldId", "p:sldMasterId"]) {
    for (const { node } of findAll(doc, tag)) {
      const id = Number(attr(node, "id"));
      if (Number.isFinite(id) && id > max) max = id;
    }
  }
  return max;
}

/** Return the children array of an ID-list element, creating the element (in schema order) if absent. */
function ensureIdList(doc: XmlDocument, listTag: string): XmlDocument {
  const existing = findAll(doc, listTag)[0];
  if (existing) return existing.children;

  const presentation = findAll(doc, "p:presentation")[0];
  if (!presentation) return []; // malformed; caller simply won't backfill
  const siblings = presentation.children;
  const newNode: XmlNode = { [listTag]: [] } as XmlNode;

  // Schema order: sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, sldIdLst, sldSz, ...
  const order = ["p:sldMasterIdLst", "p:notesMasterIdLst", "p:handoutMasterIdLst", "p:sldIdLst", "p:sldSz"];
  const myPos = order.indexOf(listTag);
  let insertAt = siblings.length;
  for (let i = 0; i < siblings.length; i++) {
    const tag = Object.keys(siblings[i]).find((k) => k !== ":@");
    const pos = tag ? order.indexOf(tag) : -1;
    if (pos > myPos) { insertAt = i; break; }
  }
  siblings.splice(insertAt, 0, newNode);
  return newNode[listTag] as XmlDocument;
}

function removeIdListEntries(pkg: OpcPackage, rIds: Set<string>, changes: Change[]): void {
  const doc = getPartXml(pkg, PRESENTATION_PART);
  if (!doc) return;
  let mutated = false;
  for (const listTag of ["p:sldIdLst", "p:sldMasterIdLst", "p:notesMasterIdLst"]) {
    const list = findAll(doc, listTag)[0];
    if (!list) continue;
    const itemTag = listTag.replace("Lst", "").replace("p:sld", "p:sldId").replace("p:notesMaster", "p:notesMasterId");
    void itemTag;
    for (let i = list.children.length - 1; i >= 0; i--) {
      const child = list.children[i];
      const bag = (child as Record<string, unknown>)[":@"] as Record<string, string> | undefined;
      const rId = bag?.["@_r:id"];
      if (rId && rIds.has(rId)) {
        list.children.splice(i, 1);
        mutated = true;
        changes.push({ ruleId: "presentationIdLists", description: `Removed ID-list entry for deleted relationship ${rId}` });
      }
    }
  }
  if (mutated) setPartXml(pkg, PRESENTATION_PART, doc);
}

// --- orchestration ----------------------------------------------------------

export interface SanitizeResult {
  bytes: Uint8Array;
  changes: Change[];
  remainingViolations: Violation[];
}

const FIXERS = [
  fixDanglingOverrides,
  fixMissingOverrides,
  fixPlatformSpecificParts,
  fixBloatedMedia,
  fixOrphanedMedia,
  fixDanglingRelationships,
  fixPresentationIdLists,
  fixNegativeExtents,
];

/** Sanitize a package that is already loaded (mutates it) and return the change log. */
export function sanitizePackage(pkg: OpcPackage): Change[] {
  const changes: Change[] = [];
  for (const fixer of FIXERS) changes.push(...fixer(pkg));
  // Drop content-type entries for the bookkeeping parts if they crept in.
  pkg.contentTypes = pkg.contentTypes.filter(
    (e) => !(e.kind === "override" && (e.partName === `/${CONTENT_TYPES_PART}` || isRelsPart(e.partName)))
  );
  return changes;
}

export async function sanitizeBuffer(bytes: Uint8Array | Buffer): Promise<SanitizeResult> {
  const pkg = await loadPackage(bytes);
  const changes = sanitizePackage(pkg);
  const outBytes = await savePackage(pkg);
  const reloaded = await loadPackage(outBytes);
  const remainingViolations = runAllRules(reloaded);
  return { bytes: outBytes, changes, remainingViolations };
}
