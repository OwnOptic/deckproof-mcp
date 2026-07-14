/**
 * Thin, fixed-options wrappers around fast-xml-parser, centralized once so
 * every rule/sanitizer parses and rebuilds XML parts identically.
 *
 * `preserveOrder: true` keeps element order intact on rebuild - many OOXML
 * schemas are order-sensitive sequences, not unordered sets, so losing order
 * would itself introduce spec violations. Attribute/tag values are left as
 * raw strings (no number/boolean coercion, no whitespace trimming) since
 * OOXML relationship IDs, hex colors, and text-run whitespace are all
 * semantically exact strings.
 */

import { XMLParser, XMLBuilder } from "fast-xml-parser";

const ATTR_PREFIX = "@_";
const TEXT_KEY = "#text";

const parserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: ATTR_PREFIX,
  textNodeName: TEXT_KEY,
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  ignoreDeclaration: false,
  // Security: these parts come from untrusted uploaded .pptx files. Disable
  // custom entity processing so a maliciously crafted part cannot trigger
  // entity-expansion ("billion laughs") amplification. OOXML does not rely on
  // custom DTD entities, so this is safe.
  processEntities: false,
};

const builderOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: ATTR_PREFIX,
  textNodeName: TEXT_KEY,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

/** An XML document parsed with preserveOrder: an ordered array of single-key element nodes. */
export type XmlNode = Record<string, unknown>;
export type XmlDocument = XmlNode[];

const parser = new XMLParser(parserOptions);
const builder = new XMLBuilder(builderOptions);

export function parseXml(xml: string | Uint8Array): XmlDocument {
  return parser.parse(xml) as XmlDocument;
}

export function buildXml(doc: XmlDocument): string {
  return builder.build(doc) as string;
}

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

export function parseXmlBytes(bytes: Uint8Array): XmlDocument {
  return parseXml(decoder.decode(bytes));
}

export function buildXmlBytes(doc: XmlDocument): Uint8Array {
  return encoder.encode(buildXml(doc));
}

/**
 * Find every node in a parsed document (at any depth) whose tag name matches
 * `tagName`, returning `{ node, attrs }` pairs. A minimal, dependency-free
 * tree walker - good enough for the small, known-shape OOXML parts this
 * engine inspects (presentation.xml, .rels files, slide XML), not a general
 * XPath engine.
 */
export function findAll(
  doc: XmlDocument,
  tagName: string
): Array<{ node: XmlNode; children: XmlDocument }> {
  const results: Array<{ node: XmlNode; children: XmlDocument }> = [];

  function walk(nodes: XmlDocument) {
    for (const node of nodes) {
      for (const key of Object.keys(node)) {
        if (key === ATTR_PREFIX.slice(0, 0)) continue; // no-op guard, keys are tag names or ":@"
        if (key === TEXT_KEY || key === ":@") continue;
        if (key === tagName) {
          results.push({ node, children: node[key] as XmlDocument });
        }
        const value = node[key];
        if (Array.isArray(value)) {
          walk(value as XmlDocument);
        }
      }
    }
  }

  walk(doc);
  return results;
}

/** Read the attribute bag (`:@`) of a preserveOrder node, or `{}` if it has none. */
export function attrsOf(node: XmlNode): Record<string, string> {
  const bag = (node as Record<string, unknown>)[":@"] as
    | Record<string, string>
    | undefined;
  return bag ?? {};
}

export function attr(node: XmlNode, name: string): string | undefined {
  return attrsOf(node)[`${ATTR_PREFIX}${name}`];
}
