/**
 * Extracts the color and font scheme from an uploaded template's
 * ppt/theme/theme1.xml, so content we generate on top of a cloned stencil
 * (the "generate on stencil" path for graphically-rich archetypes) is styled
 * in the company's own brand rather than a generic palette.
 *
 * Best-effort: any missing slot falls back to a neutral Office-like default,
 * so a malformed or unusual theme never breaks creation - it just yields a
 * plainer result.
 */

import type { OpcPackage } from "./types.js";
import { getPartXml, listContentParts } from "./opcPackage.js";
import { findAll, attr, type XmlDocument } from "./xml.js";

export interface ThemeInfo {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  majorFont: string;
  minorFont: string;
}

export const NEUTRAL_THEME: ThemeInfo = {
  dk1: "111827",
  lt1: "FFFFFF",
  dk2: "374151",
  lt2: "E7E6E6",
  accent1: "4472C4",
  accent2: "ED7D31",
  accent3: "A5A5A5",
  accent4: "FFC000",
  accent5: "5B9BD5",
  accent6: "70AD47",
  majorFont: "Calibri Light",
  minorFont: "Calibri",
};

/** Given a color-container node's children (e.g. inside <a:accent1>), pull the hex value. */
function colorFromContainer(children: XmlDocument): string | undefined {
  for (const child of children) {
    if ((child as Record<string, unknown>)["a:srgbClr"] !== undefined) {
      return attr(child, "val");
    }
    if ((child as Record<string, unknown>)["a:sysClr"] !== undefined) {
      return attr(child, "lastClr") ?? attr(child, "val");
    }
  }
  return undefined;
}

function childByTag(children: XmlDocument, tag: string): XmlDocument | undefined {
  const node = children.find((c) => (c as Record<string, unknown>)[tag] !== undefined);
  return node ? ((node as Record<string, unknown>)[tag] as XmlDocument) : undefined;
}

/** Read the `typeface` of the <a:latin> inside a font container (a:majorFont / a:minorFont). */
function latinTypefaceIn(doc: XmlDocument, fontContainerTag: string): string | undefined {
  const container = findAll(doc, fontContainerTag)[0];
  if (!container) return undefined;
  const latin = findAll(container.children, "a:latin")[0];
  return latin ? attr(latin.node, "typeface") : undefined;
}

/** Find the first theme part in the package (usually ppt/theme/theme1.xml). */
export function findThemePart(pkg: OpcPackage): string | undefined {
  return listContentParts(pkg)
    .filter((p) => p.startsWith("ppt/theme/") && p.endsWith(".xml"))
    .sort()[0];
}

export function extractTheme(pkg: OpcPackage): ThemeInfo {
  const themePart = findThemePart(pkg);
  if (!themePart) return { ...NEUTRAL_THEME };
  const doc = getPartXml(pkg, themePart);
  if (!doc) return { ...NEUTRAL_THEME };

  const result: ThemeInfo = { ...NEUTRAL_THEME };

  const clrScheme = findAll(doc, "a:clrScheme")[0];
  if (clrScheme) {
    const colorSlots: Array<keyof ThemeInfo> = [
      "dk1", "lt1", "dk2", "lt2",
      "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    ];
    for (const slot of colorSlots) {
      const container = childByTag(clrScheme.children, `a:${slot}`);
      const hex = container && colorFromContainer(container);
      if (hex) result[slot] = hex.replace(/^#/, "").toUpperCase();
    }
  }

  const majorFont = latinTypefaceIn(doc, "a:majorFont");
  if (majorFont) result.majorFont = majorFont;
  const minorFont = latinTypefaceIn(doc, "a:minorFont");
  if (minorFont) result.minorFont = minorFont;

  return result;
}
