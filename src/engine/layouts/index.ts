/**
 * Layout archetype registry (all 15). Each entry has a from-scratch PptxGenJS
 * composer plus, for template mode, either a `refillStencil` hook (text
 * archetypes) or a `generateOnStencil` hook (rich archetypes). Adding an
 * archetype means adding its schema/draw in the *Layouts.ts files and one
 * entry here.
 */

import { z } from "zod";
import type { ISlide } from "pptx-automizer";
import type { PptxSlide } from "../pptxgenTypes.js";
import type { ThemeInfo } from "../themeExtract.js";
import type { TemplateSlideInventory } from "../templateInspect.js";
import { NEUTRAL_PALETTE, paletteFromTheme, type Palette } from "./_shared.js";

import { coverPropsSchema, composeCoverFromScratch, refillCoverStencil, type CoverProps } from "./cover.js";
import * as T from "./textLayouts.js";
import * as R from "./richLayouts.js";

export interface LayoutDefinition<Props = unknown> {
  archetype: string;
  description: string;
  schema: z.ZodType<Props>;
  composeFromScratch: (slide: PptxSlide, props: Props) => void;
  refillStencil?: (slide: ISlide, props: Props, stencil: TemplateSlideInventory, theme: ThemeInfo) => void;
  generateOnStencil?: (slide: PptxSlide, props: Props, theme: ThemeInfo) => void;
}

/** Wrap a `draw(slide, props, palette)` into from-scratch (neutral) + generateOnStencil (theme). */
function rich<P>(
  archetype: string,
  description: string,
  schema: z.ZodType<P>,
  draw: (slide: PptxSlide, props: P, p: Palette) => void
): LayoutDefinition<P> {
  return {
    archetype, description, schema,
    composeFromScratch: (slide, props) => draw(slide, props, NEUTRAL_PALETTE),
    generateOnStencil: (slide, props, theme) => draw(slide, props, paletteFromTheme(theme)),
  };
}

/** Wrap a `draw` + `refill` into from-scratch (neutral) + refillStencil. */
function text<P>(
  archetype: string,
  description: string,
  schema: z.ZodType<P>,
  draw: (slide: PptxSlide, props: P, p: Palette) => void,
  refillFn: (slide: ISlide, props: P, stencil: TemplateSlideInventory) => void
): LayoutDefinition<P> {
  return {
    archetype, description, schema,
    composeFromScratch: (slide, props) => draw(slide, props, NEUTRAL_PALETTE),
    refillStencil: (slide, props, stencil) => refillFn(slide, props, stencil),
  };
}

const defs: Array<LayoutDefinition<any>> = [
  {
    archetype: "cover",
    description: "Title slide: eyebrow + title + subtitle + optional context line.",
    schema: coverPropsSchema,
    composeFromScratch: composeCoverFromScratch as (s: PptxSlide, p: unknown) => void,
    refillStencil: refillCoverStencil as LayoutDefinition["refillStencil"],
  },
  text("agenda", "Agenda / table of contents with optional durations.", T.agendaSchema, T.drawAgenda, T.refillAgenda),
  text("contentBullets", "Title, optional intro, and a bulleted list.", T.contentBulletsSchema, T.drawContentBullets, T.refillContentBullets),
  text("twoColumns", "Two side-by-side columns of headings/bullets/body.", T.twoColumnsSchema, T.drawTwoColumns, T.refillTwoColumns),
  text("quote", "A large centered pull quote with attribution.", T.quoteSchema, T.drawQuote, T.refillQuote),
  text("sectionDivider", "Section/chapter divider with a title and optional subtitle.", T.sectionDividerSchema, T.drawSectionDivider, T.refillSectionDivider),
  text("closingNextSteps", "Closing slide listing next steps with owners.", T.closingNextStepsSchema, T.drawClosingNextSteps, T.refillClosingNextSteps),
  rich("comparisonTable", "Compare options across features (checkmarks / values).", R.comparisonTableSchema, R.drawComparisonTable),
  rich("dataTable", "A plain data table with headers and rows.", R.dataTableSchema, R.drawDataTable),
  rich("timeline", "A horizontal timeline of milestones.", R.timelineSchema, R.drawTimeline),
  rich("statsBanner", "A row of big KPI numbers with labels.", R.statsBannerSchema, R.drawStatsBanner),
  rich("cardGrid", "A grid of titled cards.", R.cardGridSchema, R.drawCardGrid),
  rich("orgChart", "A one-level org chart (root + children).", R.orgChartSchema, R.drawOrgChart),
  rich("matrixQuadrant", "A 2x2 matrix with labeled axes and four quadrants.", R.matrixQuadrantSchema, R.drawMatrixQuadrant),
  rich("verticalSteps", "A numbered vertical process/step list.", R.verticalStepsSchema, R.drawVerticalSteps),
];

export const layoutRegistry: Record<string, LayoutDefinition<any>> =
  Object.fromEntries(defs.map((d) => [d.archetype, d]));

/** Discriminated union over every archetype's props schema (keyed on `archetype`). */
export const slideSpecSchema = z.discriminatedUnion("archetype", [
  coverPropsSchema,
  T.agendaSchema, T.contentBulletsSchema, T.twoColumnsSchema, T.quoteSchema, T.sectionDividerSchema, T.closingNextStepsSchema,
  R.comparisonTableSchema, R.dataTableSchema, R.timelineSchema, R.statsBannerSchema, R.cardGridSchema, R.orgChartSchema, R.matrixQuadrantSchema, R.verticalStepsSchema,
]);

export type SlideSpec = z.infer<typeof slideSpecSchema>;
export type { CoverProps };

/**
 * Per-slide input for pptx_create_deck: archetype props plus an optional
 * `stencilSlideIndex` (template mode). Intersection keeps it valid over the
 * discriminated union.
 */
export const slideInputSchema = z.intersection(
  slideSpecSchema,
  z.object({
    stencilSlideIndex: z
      .number()
      .int()
      .min(0)
      .optional()
      .describe("Template mode only: 0-based index of the template slide to clone as this slide's stencil."),
  })
);

export type SlideInput = z.infer<typeof slideInputSchema>;

export function listLayouts(): Array<{ archetype: string; description: string }> {
  return defs.map((d) => ({ archetype: d.archetype, description: d.description }));
}
