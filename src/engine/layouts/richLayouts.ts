/**
 * Graphically-rich archetypes: comparisonTable, dataTable, timeline,
 * statsBanner, cardGrid, orgChart, matrixQuadrant, verticalSteps. Each has one
 * `draw(slide, props, palette)` used by both the from-scratch composer
 * (neutral palette) and the `generateOnStencil` hook (palette from the
 * template's theme, so content generated on top of a cloned stencil stays on
 * brand). All shapes route through resolvePreset to avoid invalid `prst`.
 */

import { z } from "zod";
import type { PptxSlide } from "../pptxgenTypes.js";
import { resolvePreset } from "../presetGeometry.js";
import { MARGIN, CONTENT_W, SLIDE_W, columns, drawTitle, type Palette } from "./_shared.js";

const rect = resolvePreset("rect");
const roundRect = resolvePreset("roundRect");
const ellipse = resolvePreset("ellipse");

// --- comparisonTable --------------------------------------------------------

export const comparisonTableSchema = z.object({
  archetype: z.literal("comparisonTable"),
  title: z.string(),
  options: z.array(z.string()).min(2).describe("Column headers (the things being compared)."),
  features: z.array(z.object({
    name: z.string(),
    values: z.array(z.union([z.boolean(), z.string()])),
  })).min(1),
});
export type ComparisonTableProps = z.infer<typeof comparisonTableSchema>;

export function drawComparisonTable(slide: PptxSlide, props: ComparisonTableProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const header = [{ text: "", options: { fill: { color: p.band } } },
    ...props.options.map((o) => ({ text: o, options: { bold: true, color: "FFFFFF", fill: { color: p.accent }, align: "center" } }))];
  const body = props.features.map((f) => [
    { text: f.name, options: { bold: true, color: p.text } },
    ...f.values.map((v) => ({
      text: typeof v === "boolean" ? (v ? "✓" : "—") : v,
      options: { align: "center", color: typeof v === "boolean" ? (v ? p.accent : p.subtle) : p.text },
    })),
  ]);
  slide.addTable([header, ...body], { x: MARGIN, y: 1.7, w: CONTENT_W, fontSize: 13, border: { type: "solid", color: p.band, pt: 1 }, valign: "middle" });
}

// --- dataTable --------------------------------------------------------------

export const dataTableSchema = z.object({
  archetype: z.literal("dataTable"),
  title: z.string(),
  headers: z.array(z.string()).min(1),
  rows: z.array(z.array(z.string())).min(1),
});
export type DataTableProps = z.infer<typeof dataTableSchema>;

export function drawDataTable(slide: PptxSlide, props: DataTableProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const header = props.headers.map((h) => ({ text: h, options: { bold: true, color: "FFFFFF", fill: { color: p.accent } } }));
  const body = props.rows.map((r) => r.map((c) => ({ text: c, options: { color: p.text } })));
  slide.addTable([header, ...body], { x: MARGIN, y: 1.7, w: CONTENT_W, fontSize: 13, border: { type: "solid", color: p.band, pt: 1 }, valign: "middle" });
}

// --- timeline ---------------------------------------------------------------

export const timelineSchema = z.object({
  archetype: z.literal("timeline"),
  title: z.string(),
  milestones: z.array(z.object({ when: z.string(), label: z.string(), detail: z.string().optional() })).min(1),
});
export type TimelineProps = z.infer<typeof timelineSchema>;

export function drawTimeline(slide: PptxSlide, props: TimelineProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const y = 3.6;
  slide.addShape(rect, { x: MARGIN, y, w: CONTENT_W, h: 0.06, fill: { color: p.accent } });
  const cols = columns(props.milestones.length, 0.2);
  props.milestones.forEach((m, i) => {
    const cx = cols[i].x + cols[i].w / 2;
    slide.addShape(ellipse, { x: cx - 0.1, y: y - 0.07, w: 0.2, h: 0.2, fill: { color: p.accent } });
    slide.addText(m.when, { x: cols[i].x, y: y - 0.9, w: cols[i].w, h: 0.4, align: "center", bold: true, fontSize: 13, color: p.accent });
    slide.addText([{ text: m.label, options: { bold: true, fontSize: 13, color: p.text } },
      ...(m.detail ? [{ text: `\n${m.detail}`, options: { fontSize: 11, color: p.subtle } }] : [])],
      { x: cols[i].x, y: y + 0.25, w: cols[i].w, h: 1.4, align: "center" });
  });
}

// --- statsBanner ------------------------------------------------------------

export const statsBannerSchema = z.object({
  archetype: z.literal("statsBanner"),
  title: z.string().optional(),
  stats: z.array(z.object({ value: z.string(), label: z.string(), trend: z.string().optional() })).min(1),
});
export type StatsBannerProps = z.infer<typeof statsBannerSchema>;

export function drawStatsBanner(slide: PptxSlide, props: StatsBannerProps, p: Palette): void {
  if (props.title) drawTitle(slide, props.title, p);
  const y = props.title ? 2.6 : 2.9;
  const cols = columns(props.stats.length, 0.3);
  props.stats.forEach((s, i) => {
    slide.addText(s.value, { x: cols[i].x, y, w: cols[i].w, h: 1.1, align: "center", fontSize: 44, bold: true, color: p.accent });
    slide.addText(s.label, { x: cols[i].x, y: y + 1.15, w: cols[i].w, h: 0.6, align: "center", fontSize: 14, color: p.text });
    if (s.trend) slide.addText(s.trend, { x: cols[i].x, y: y + 1.7, w: cols[i].w, h: 0.4, align: "center", fontSize: 12, color: p.subtle });
  });
}

// --- cardGrid ---------------------------------------------------------------

export const cardGridSchema = z.object({
  archetype: z.literal("cardGrid"),
  title: z.string(),
  cards: z.array(z.object({ title: z.string(), body: z.string().optional() })).min(1),
});
export type CardGridProps = z.infer<typeof cardGridSchema>;

export function drawCardGrid(slide: PptxSlide, props: CardGridProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const perRow = props.cards.length <= 3 ? props.cards.length : props.cards.length <= 4 ? 2 : 3;
  const cols = columns(perRow, 0.3);
  const cardH = 1.7, gapY = 0.3;
  props.cards.forEach((c, i) => {
    const col = cols[i % perRow];
    const row = Math.floor(i / perRow);
    const y = 1.7 + row * (cardH + gapY);
    slide.addShape(roundRect, { x: col.x, y, w: col.w, h: cardH, fill: { color: p.band }, line: { color: p.band } });
    slide.addText(c.title, { x: col.x + 0.2, y: y + 0.15, w: col.w - 0.4, h: 0.5, bold: true, fontSize: 15, color: p.accent });
    if (c.body) slide.addText(c.body, { x: col.x + 0.2, y: y + 0.7, w: col.w - 0.4, h: cardH - 0.85, fontSize: 12, color: p.text, valign: "top" });
  });
}

// --- orgChart ---------------------------------------------------------------

export const orgChartSchema = z.object({
  archetype: z.literal("orgChart"),
  title: z.string(),
  root: z.object({ title: z.string(), subtitle: z.string().optional() }),
  children: z.array(z.object({ title: z.string(), subtitle: z.string().optional() })).min(1),
});
export type OrgChartProps = z.infer<typeof orgChartSchema>;

function drawBox(slide: PptxSlide, x: number, y: number, w: number, h: number, title: string, subtitle: string | undefined, p: Palette, filled: boolean): void {
  slide.addShape(roundRect, { x, y, w, h, fill: { color: filled ? p.accent : p.band }, line: { color: p.accent } });
  slide.addText([{ text: title, options: { bold: true, fontSize: 14, color: filled ? "FFFFFF" : p.text } },
    ...(subtitle ? [{ text: `\n${subtitle}`, options: { fontSize: 11, color: filled ? "FFFFFF" : p.subtle } }] : [])],
    { x, y, w, h, align: "center", valign: "middle" });
}

export function drawOrgChart(slide: PptxSlide, props: OrgChartProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const rootW = 3, rootX = (SLIDE_W - rootW) / 2, rootY = 1.7;
  drawBox(slide, rootX, rootY, rootW, 0.9, props.root.title, props.root.subtitle, p, true);
  const cols = columns(Math.min(props.children.length, 4), 0.3);
  const childY = 3.6;
  props.children.slice(0, 4).forEach((c, i) => {
    slide.addShape(rect, { x: cols[i].x + cols[i].w / 2 - 0.01, y: rootY + 0.9, w: 0.02, h: childY - (rootY + 0.9), fill: { color: p.accent } });
    drawBox(slide, cols[i].x, childY, cols[i].w, 0.9, c.title, c.subtitle, p, false);
  });
}

// --- matrixQuadrant ---------------------------------------------------------

export const matrixQuadrantSchema = z.object({
  archetype: z.literal("matrixQuadrant"),
  title: z.string(),
  xAxis: z.string().optional(),
  yAxis: z.string().optional(),
  quadrants: z.array(z.object({ heading: z.string(), items: z.array(z.string()).optional() })).length(4),
});
export type MatrixQuadrantProps = z.infer<typeof matrixQuadrantSchema>;

export function drawMatrixQuadrant(slide: PptxSlide, props: MatrixQuadrantProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const gx = MARGIN + 0.4, gy = 1.8, gw = CONTENT_W - 0.4, gh = 4.6;
  const halfW = gw / 2, halfH = gh / 2;
  const cells = [ [gx, gy], [gx + halfW, gy], [gx, gy + halfH], [gx + halfW, gy + halfH] ];
  props.quadrants.forEach((q, i) => {
    const [x, y] = cells[i];
    slide.addShape(rect, { x, y, w: halfW - 0.05, h: halfH - 0.05, fill: { color: i % 3 === 0 ? p.band : "FFFFFF" }, line: { color: p.accent, pt: 1 } });
    slide.addText([{ text: q.heading, options: { bold: true, fontSize: 14, color: p.accent } },
      ...(q.items?.length ? [{ text: `\n${q.items.join("\n")}`, options: { fontSize: 12, color: p.text } }] : [])],
      { x: x + 0.15, y: y + 0.15, w: halfW - 0.35, h: halfH - 0.35, valign: "top" });
  });
  if (props.xAxis) slide.addText(props.xAxis, { x: gx, y: gy + gh + 0.05, w: gw, h: 0.4, align: "center", fontSize: 12, italic: true, color: p.subtle });
  if (props.yAxis) slide.addText(props.yAxis, { x: MARGIN - 0.4, y: gy, w: gh, h: 0.4, align: "center", fontSize: 12, italic: true, color: p.subtle, rotate: 270 });
}

// --- verticalSteps ----------------------------------------------------------

export const verticalStepsSchema = z.object({
  archetype: z.literal("verticalSteps"),
  title: z.string(),
  steps: z.array(z.object({ title: z.string(), detail: z.string().optional() })).min(1),
});
export type VerticalStepsProps = z.infer<typeof verticalStepsSchema>;

export function drawVerticalSteps(slide: PptxSlide, props: VerticalStepsProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  let y = 1.8;
  const h = Math.min(1.0, (5.2) / props.steps.length);
  props.steps.forEach((s, i) => {
    slide.addShape(ellipse, { x: MARGIN, y, w: 0.55, h: 0.55, fill: { color: p.accent } });
    slide.addText(`${i + 1}`, { x: MARGIN, y, w: 0.55, h: 0.55, align: "center", valign: "middle", bold: true, color: "FFFFFF", fontSize: 16 });
    slide.addText([{ text: s.title, options: { bold: true, fontSize: 16, color: p.text } },
      ...(s.detail ? [{ text: `  ${s.detail}`, options: { fontSize: 13, color: p.subtle } }] : [])],
      { x: MARGIN + 0.75, y, w: CONTENT_W - 0.75, h: 0.6, valign: "middle" });
    y += h + 0.15;
  });
}
