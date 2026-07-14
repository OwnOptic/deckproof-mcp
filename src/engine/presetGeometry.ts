/**
 * ECMA-376 DrawingML preset shape geometries (the valid values for a shape's
 * `prst`), plus a friendly-alias map. PptxGenJS passes shape-type strings
 * straight through without validating them, so an invalid preset (e.g. "oval"
 * instead of "ellipse") silently produces a shape PowerPoint strips on repair.
 * Every shape-emitting layout resolves its shape name through `resolvePreset`.
 */

const ALIASES: Record<string, string> = {
  oval: "ellipse",
  circle: "ellipse",
  rect: "rect",
  rectangle: "rect",
  roundedrect: "roundRect",
  "rounded-rectangle": "roundRect",
  square: "rect",
  arrow: "rightArrow",
  triangle: "triangle",
  diamond: "diamond",
  line: "line",
};

/** A subset of the ECMA-376 preset list - the ones the built-in layouts use. */
const KNOWN_PRESETS = new Set([
  "rect", "roundRect", "ellipse", "triangle", "diamond", "rightArrow",
  "leftArrow", "chevron", "pentagon", "hexagon", "line", "plus",
  "flowChartProcess", "flowChartDecision", "star5", "wedgeRectCallout",
]);

export type ShapeName = string;

/**
 * Resolve a friendly or exact shape name to a valid ECMA-376 preset.
 * Unknown names fall back to "rect" (a universally valid preset) rather than
 * risking an invalid `prst` that PowerPoint would strip.
 */
export function resolvePreset(name: ShapeName): string {
  const lower = name.toLowerCase();
  if (KNOWN_PRESETS.has(name)) return name;
  if (ALIASES[lower]) return ALIASES[lower];
  if (KNOWN_PRESETS.has(lower)) return lower;
  return "rect";
}
