/**
 * Shared helpers for layout composers. A layout draws the same way whether
 * building from scratch (neutral palette) or generating on top of a cloned
 * stencil (palette derived from the template's theme), so each layout exposes
 * a single `draw(slide, props, palette)` and the two hooks just pass a
 * different palette.
 */

import type { PptxSlide } from "../pptxgenTypes.js";
import type { ThemeInfo } from "../themeExtract.js";

/** Widescreen 13.33 x 7.5in canvas constants. */
export const SLIDE_W = 13.33;
export const SLIDE_H = 7.5;
export const MARGIN = 0.6;
export const CONTENT_W = SLIDE_W - MARGIN * 2;

export interface Palette {
  text: string;
  subtle: string;
  accent: string;
  accentAlt: string;
  bg: string;
  band: string;
}

export const NEUTRAL_PALETTE: Palette = {
  text: "111827",
  subtle: "6B7280",
  accent: "2563EB",
  accentAlt: "7C3AED",
  bg: "FFFFFF",
  band: "F3F4F6",
};

export function paletteFromTheme(theme: ThemeInfo): Palette {
  return {
    text: theme.dk1,
    subtle: theme.dk2,
    accent: theme.accent1,
    accentAlt: theme.accent2,
    bg: theme.lt1,
    band: theme.lt2,
  };
}

/** Standard slide title, placed consistently across layouts. */
export function drawTitle(slide: PptxSlide, title: string, p: Palette): void {
  slide.addText(title, {
    x: MARGIN, y: 0.45, w: CONTENT_W, h: 0.9,
    fontSize: 28, bold: true, color: p.text,
  });
}

/** Even column x-positions for `n` columns across the content width, with a gap. */
export function columns(n: number, gap = 0.3): Array<{ x: number; w: number }> {
  const w = (CONTENT_W - gap * (n - 1)) / n;
  return Array.from({ length: n }, (_, i) => ({ x: MARGIN + i * (w + gap), w }));
}
