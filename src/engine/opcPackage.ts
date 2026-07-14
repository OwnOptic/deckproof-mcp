/**
 * Open Packaging Conventions (OPC) package abstraction - the single layer
 * every rule, the sanitizer, and the creators use to read/write a .pptx.
 * No rule ever touches raw zip bytes or relationship/content-type XML
 * directly; they go through the typed helpers here instead.
 */

import JSZip from "jszip";
import type {
  OpcPackage,
  OpcRelationship,
  ContentTypeEntry,
} from "./types.js";
import { findAll, attrsOf, buildXml, parseXml, type XmlDocument } from "./xml.js";

const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
const CONTENT_TYPES_PART = "[Content_Types].xml";

const textDecoder = new TextDecoder("utf-8");
const textEncoder = new TextEncoder();

/** Given a part name, return the path of its .rels file (root part -> "" -> "_rels/.rels"). */
export function relsPathFor(partName: string): string {
  if (partName === "") return "_rels/.rels";
  const lastSlash = partName.lastIndexOf("/");
  const dir = lastSlash === -1 ? "" : partName.slice(0, lastSlash);
  const file = lastSlash === -1 ? partName : partName.slice(lastSlash + 1);
  return dir === "" ? `_rels/${file}.rels` : `${dir}/_rels/${file}.rels`;
}

/** Inverse of relsPathFor: given a .rels part path, return the part it describes. */
export function sourcePartForRelsPath(relsPath: string): string {
  const withoutRels = relsPath.replace(/\.rels$/, "");
  return withoutRels.replace(/(^|\/)_rels\//, "$1");
}

export function isRelsPart(partName: string): boolean {
  return partName.endsWith(".rels");
}

// ---------------------------------------------------------------------------
// Relationship XML parse/build
// ---------------------------------------------------------------------------

export function parseRelationshipsXml(xml: string): OpcRelationship[] {
  const doc = parseXml(xml);
  const rels: OpcRelationship[] = [];
  for (const { node } of findAll(doc, "Relationship")) {
    const attrs = attrsOf(node);
    const targetMode = attrs["@_TargetMode"];
    rels.push({
      id: attrs["@_Id"] ?? "",
      type: attrs["@_Type"] ?? "",
      target: attrs["@_Target"] ?? "",
      ...(targetMode === "External" || targetMode === "Internal"
        ? { targetMode }
        : {}),
    });
  }
  return rels;
}

export function buildRelationshipsXml(rels: OpcRelationship[]): string {
  const doc: XmlDocument = [
    { "?xml": [], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } },
    {
      Relationships: rels.map((r) => ({
        Relationship: [],
        ":@": {
          "@_Id": r.id,
          "@_Type": r.type,
          "@_Target": r.target,
          ...(r.targetMode ? { "@_TargetMode": r.targetMode } : {}),
        },
      })),
      ":@": { "@_xmlns": RELS_NS },
    },
  ];
  return buildXml(doc);
}

// ---------------------------------------------------------------------------
// [Content_Types].xml parse/build
// ---------------------------------------------------------------------------

export function parseContentTypesXml(xml: string): ContentTypeEntry[] {
  const doc = parseXml(xml);
  const entries: ContentTypeEntry[] = [];
  for (const { node } of findAll(doc, "Default")) {
    const attrs = attrsOf(node);
    entries.push({
      kind: "default",
      extension: attrs["@_Extension"] ?? "",
      contentType: attrs["@_ContentType"] ?? "",
    });
  }
  for (const { node } of findAll(doc, "Override")) {
    const attrs = attrsOf(node);
    entries.push({
      kind: "override",
      partName: attrs["@_PartName"] ?? "",
      contentType: attrs["@_ContentType"] ?? "",
    });
  }
  return entries;
}

export function buildContentTypesXml(entries: ContentTypeEntry[]): string {
  const doc: XmlDocument = [
    { "?xml": [], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } },
    {
      Types: entries.map((e) =>
        e.kind === "default"
          ? {
              Default: [],
              ":@": { "@_Extension": e.extension, "@_ContentType": e.contentType },
            }
          : {
              Override: [],
              ":@": { "@_PartName": e.partName, "@_ContentType": e.contentType },
            }
      ),
      ":@": { "@_xmlns": CONTENT_TYPES_NS },
    },
  ];
  return buildXml(doc);
}

// ---------------------------------------------------------------------------
// Package load / save
// ---------------------------------------------------------------------------

/** Max total uncompressed size when loading a package, to bound zip-bomb amplification. */
export const MAX_UNCOMPRESSED_BYTES = 300 * 1024 * 1024; // 300 MB

export async function loadPackage(
  bytes: Uint8Array | Buffer
): Promise<OpcPackage> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (err) {
    throw new Error(
      `Not a valid .pptx (Open Packaging Conventions zip): ${err instanceof Error ? err.message : String(err)}`
    );
  }
  const parts = new Map<string, Uint8Array>();
  const relsByPart = new Map<string, OpcRelationship[]>();
  let contentTypes: ContentTypeEntry[] = [];

  const entries = Object.values(zip.files).filter((f) => !f.dir);
  let totalBytes = 0;
  for (const entry of entries) {
    const data = await entry.async("uint8array");
    totalBytes += data.byteLength;
    if (totalBytes > MAX_UNCOMPRESSED_BYTES) {
      throw new Error(
        `Refusing to load: uncompressed package exceeds ${Math.round(MAX_UNCOMPRESSED_BYTES / (1024 * 1024))} MB (possible zip bomb).`
      );
    }
    parts.set(entry.name, data);
  }

  for (const [partName, data] of parts) {
    if (partName === CONTENT_TYPES_PART) {
      contentTypes = parseContentTypesXml(textDecoder.decode(data));
    } else if (isRelsPart(partName)) {
      const source = sourcePartForRelsPath(partName);
      relsByPart.set(source, parseRelationshipsXml(textDecoder.decode(data)));
    }
  }

  return { parts, relsByPart, contentTypes };
}

export async function savePackage(pkg: OpcPackage): Promise<Uint8Array> {
  const zip = new JSZip();

  for (const [partName, data] of pkg.parts) {
    if (partName === CONTENT_TYPES_PART || isRelsPart(partName)) continue;
    zip.file(partName, data);
  }

  zip.file(CONTENT_TYPES_PART, buildContentTypesXml(pkg.contentTypes));

  for (const [source, rels] of pkg.relsByPart) {
    if (rels.length === 0) continue;
    zip.file(relsPathFor(source), buildRelationshipsXml(rels));
  }

  return zip.generateAsync({ type: "uint8array" });
}

// ---------------------------------------------------------------------------
// Part-level convenience helpers
// ---------------------------------------------------------------------------

export function hasPart(pkg: OpcPackage, partName: string): boolean {
  return pkg.parts.has(partName);
}

export function getPartText(pkg: OpcPackage, partName: string): string | undefined {
  const bytes = pkg.parts.get(partName);
  return bytes ? textDecoder.decode(bytes) : undefined;
}

export function setPartText(pkg: OpcPackage, partName: string, text: string): void {
  pkg.parts.set(partName, textEncoder.encode(text));
}

export function getPartXml(pkg: OpcPackage, partName: string): XmlDocument | undefined {
  const text = getPartText(pkg, partName);
  return text ? parseXml(text) : undefined;
}

export function setPartXml(pkg: OpcPackage, partName: string, doc: XmlDocument): void {
  setPartText(pkg, partName, buildXml(doc));
}

export function deletePart(pkg: OpcPackage, partName: string): void {
  pkg.parts.delete(partName);
  pkg.relsByPart.delete(partName);
}

/** All part names, excluding the content-types and .rels bookkeeping parts. */
export function listContentParts(pkg: OpcPackage): string[] {
  return [...pkg.parts.keys()].filter(
    (name) => name !== CONTENT_TYPES_PART && !isRelsPart(name)
  );
}

/** Content parts whose path starts with `prefix` (e.g. "ppt/slides/"). */
export function partsWithPrefix(pkg: OpcPackage, prefix: string): string[] {
  return listContentParts(pkg).filter((p) => p.startsWith(prefix));
}

/** Set of every internal (non-External) relationship target, resolved to absolute part paths. */
export function resolvedInternalTargets(pkg: OpcPackage): Set<string> {
  const targets = new Set<string>();
  for (const [sourcePart, rels] of pkg.relsByPart) {
    for (const rel of rels) {
      if (rel.targetMode === "External") continue;
      targets.add(resolveRelationshipTarget(sourcePart, rel.target));
    }
  }
  return targets;
}

/** Relationships on a part filtered to a relationship-type trailing segment (e.g. "slideLayout"). */
export function relationshipsOfType(
  pkg: OpcPackage,
  sourcePart: string,
  typeSuffix: string
): OpcRelationship[] {
  return (pkg.relsByPart.get(sourcePart) ?? []).filter(
    (r) => r.type.split("/").pop() === typeSuffix
  );
}

/**
 * Resolve a relationship's `Target` (which is relative to the *directory*
 * containing the source part, per OPC conventions) into an absolute
 * package-internal part path, normalizing "../" segments.
 */
export function resolveRelationshipTarget(sourcePart: string, target: string): string {
  const lastSlash = sourcePart.lastIndexOf("/");
  const baseDir = lastSlash === -1 ? "" : sourcePart.slice(0, lastSlash);
  const combined = baseDir === "" ? target : `${baseDir}/${target}`;

  const segments: string[] = [];
  for (const segment of combined.split("/")) {
    if (segment === "." || segment === "") continue;
    if (segment === "..") segments.pop();
    else segments.push(segment);
  }
  return segments.join("/");
}
