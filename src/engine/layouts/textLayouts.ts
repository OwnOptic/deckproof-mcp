/**
 * Text archetypes: agenda, contentBullets, twoColumns, quote, sectionDivider,
 * closingNextSteps. Each has a from-scratch composer and a `refillStencil`
 * that pours its primary text into the cloned stencil's title/body
 * placeholders (branding inherited from the stencil).
 */

import { z } from "zod";
import type { ISlide } from "pptx-automizer";
import type { PptxSlide } from "../pptxgenTypes.js";
import type { TemplateSlideInventory, TemplateElement } from "../templateInspect.js";
import { ModifyTextHelper, elementSelector } from "../automizer.js";
import { MARGIN, CONTENT_W, columns, drawTitle, NEUTRAL_PALETTE, type Palette } from "./_shared.js";

// --- helpers for stencil refill ---------------------------------------------

function findPlaceholder(stencil: TemplateSlideInventory, types: string[]): TemplateElement | undefined {
  return stencil.elements.find((e) => e.placeholderType && types.includes(e.placeholderType));
}

function refill(slide: ISlide, stencil: TemplateSlideInventory, titleText: string | undefined, bodyText: string | undefined): void {
  const titleEl = findPlaceholder(stencil, ["title", "ctrTitle"]);
  if (titleEl && titleText) slide.modifyElement(elementSelector(titleEl), ModifyTextHelper.setText(titleText));
  const bodyEl = findPlaceholder(stencil, ["body", "subTitle"]);
  if (bodyEl && bodyText) slide.modifyElement(elementSelector(bodyEl), ModifyTextHelper.setText(bodyText));
}

// --- agenda -----------------------------------------------------------------

export const agendaSchema = z.object({
  archetype: z.literal("agenda"),
  title: z.string(),
  items: z.array(z.object({
    label: z.string(),
    duration: z.string().optional(),
    detail: z.string().optional(),
  })).min(1),
});
export type AgendaProps = z.infer<typeof agendaSchema>;

export function drawAgenda(slide: PptxSlide, props: AgendaProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  let y = 1.7;
  props.items.forEach((item, i) => {
    slide.addText(`${i + 1}`, { x: MARGIN, y, w: 0.6, h: 0.6, fontSize: 20, bold: true, color: p.accent });
    slide.addText(
      [{ text: item.label, options: { bold: true, color: p.text, fontSize: 18 } },
       ...(item.detail ? [{ text: `  ${item.detail}`, options: { color: p.subtle, fontSize: 13 } }] : [])],
      { x: MARGIN + 0.7, y, w: CONTENT_W - 2, h: 0.6 }
    );
    if (item.duration) slide.addText(item.duration, { x: MARGIN + CONTENT_W - 1.4, y, w: 1.4, h: 0.6, align: "right", color: p.subtle, fontSize: 13 });
    y += 0.75;
  });
}

// --- contentBullets ---------------------------------------------------------

export const contentBulletsSchema = z.object({
  archetype: z.literal("contentBullets"),
  title: z.string(),
  intro: z.string().optional(),
  bullets: z.array(z.string()).min(1),
});
export type ContentBulletsProps = z.infer<typeof contentBulletsSchema>;

export function drawContentBullets(slide: PptxSlide, props: ContentBulletsProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  let y = 1.7;
  if (props.intro) { slide.addText(props.intro, { x: MARGIN, y, w: CONTENT_W, h: 0.6, fontSize: 16, color: p.subtle }); y += 0.8; }
  slide.addText(
    props.bullets.map((b) => ({ text: b, options: { bullet: true, color: p.text, fontSize: 16, paraSpaceAfter: 8 } })),
    { x: MARGIN, y, w: CONTENT_W, h: 5 - (y - 1.7) }
  );
}

// --- twoColumns -------------------------------------------------------------

const columnSchema = z.object({ heading: z.string().optional(), bullets: z.array(z.string()).optional(), body: z.string().optional() });
export const twoColumnsSchema = z.object({
  archetype: z.literal("twoColumns"),
  title: z.string(),
  left: columnSchema,
  right: columnSchema,
});
export type TwoColumnsProps = z.infer<typeof twoColumnsSchema>;

function drawColumn(slide: PptxSlide, col: z.infer<typeof columnSchema>, x: number, w: number, p: Palette): void {
  let y = 1.7;
  if (col.heading) { slide.addText(col.heading, { x, y, w, h: 0.6, fontSize: 18, bold: true, color: p.accent }); y += 0.75; }
  if (col.bullets?.length) {
    slide.addText(col.bullets.map((b) => ({ text: b, options: { bullet: true, color: p.text, fontSize: 15, paraSpaceAfter: 6 } })), { x, y, w, h: 4 });
  } else if (col.body) {
    slide.addText(col.body, { x, y, w, h: 4, fontSize: 15, color: p.text });
  }
}

export function drawTwoColumns(slide: PptxSlide, props: TwoColumnsProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  const [c1, c2] = columns(2);
  drawColumn(slide, props.left, c1.x, c1.w, p);
  drawColumn(slide, props.right, c2.x, c2.w, p);
}

// --- quote ------------------------------------------------------------------

export const quoteSchema = z.object({
  archetype: z.literal("quote"),
  quote: z.string(),
  author: z.string().optional(),
  role: z.string().optional(),
});
export type QuoteProps = z.infer<typeof quoteSchema>;

export function drawQuote(slide: PptxSlide, props: QuoteProps, p: Palette): void {
  slide.addText(`“${props.quote}”`, { x: MARGIN, y: 2.2, w: CONTENT_W, h: 2.5, fontSize: 30, italic: true, color: p.text, align: "center" });
  const attribution = [props.author, props.role].filter(Boolean).join(", ");
  if (attribution) slide.addText(attribution, { x: MARGIN, y: 4.9, w: CONTENT_W, h: 0.6, fontSize: 16, color: p.accent, align: "center" });
}

// --- sectionDivider ---------------------------------------------------------

export const sectionDividerSchema = z.object({
  archetype: z.literal("sectionDivider"),
  title: z.string(),
  subtitle: z.string().optional(),
});
export type SectionDividerProps = z.infer<typeof sectionDividerSchema>;

export function drawSectionDivider(slide: PptxSlide, props: SectionDividerProps, p: Palette): void {
  slide.addShape("rect", { x: 0, y: 3.2, w: 0.25, h: 1.1, fill: { color: p.accent } });
  slide.addText(props.title, { x: MARGIN, y: 3.0, w: CONTENT_W, h: 1.0, fontSize: 36, bold: true, color: p.text });
  if (props.subtitle) slide.addText(props.subtitle, { x: MARGIN, y: 4.1, w: CONTENT_W, h: 0.6, fontSize: 18, color: p.subtle });
}

// --- closingNextSteps -------------------------------------------------------

export const closingNextStepsSchema = z.object({
  archetype: z.literal("closingNextSteps"),
  title: z.string(),
  steps: z.array(z.object({ title: z.string(), note: z.string().optional(), owner: z.string().optional() })).min(1),
});
export type ClosingNextStepsProps = z.infer<typeof closingNextStepsSchema>;

export function drawClosingNextSteps(slide: PptxSlide, props: ClosingNextStepsProps, p: Palette): void {
  drawTitle(slide, props.title, p);
  let y = 1.8;
  props.steps.forEach((s, i) => {
    slide.addShape("ellipse", { x: MARGIN, y, w: 0.5, h: 0.5, fill: { color: p.accent } });
    slide.addText(`${i + 1}`, { x: MARGIN, y, w: 0.5, h: 0.5, align: "center", color: "FFFFFF", bold: true, fontSize: 14 });
    slide.addText(
      [{ text: s.title, options: { bold: true, fontSize: 16, color: p.text } },
       ...(s.note ? [{ text: `  ${s.note}`, options: { fontSize: 13, color: p.subtle } }] : [])],
      { x: MARGIN + 0.7, y, w: CONTENT_W - 2.2, h: 0.6 }
    );
    if (s.owner) slide.addText(s.owner, { x: MARGIN + CONTENT_W - 1.5, y, w: 1.5, h: 0.6, align: "right", fontSize: 13, color: p.accent });
    y += 0.8;
  });
}

// --- stencil refill wiring --------------------------------------------------

export function refillAgenda(slide: ISlide, props: AgendaProps, stencil: TemplateSlideInventory): void {
  refill(slide, stencil, props.title, props.items.map((i) => `${i.label}${i.duration ? ` (${i.duration})` : ""}`).join("\n"));
}
export function refillContentBullets(slide: ISlide, props: ContentBulletsProps, stencil: TemplateSlideInventory): void {
  refill(slide, stencil, props.title, [props.intro, ...props.bullets].filter(Boolean).join("\n"));
}
export function refillTwoColumns(slide: ISlide, props: TwoColumnsProps, stencil: TemplateSlideInventory): void {
  const body = [props.left, props.right].map((c) => [c.heading, ...(c.bullets ?? []), c.body].filter(Boolean).join("\n")).join("\n\n");
  refill(slide, stencil, props.title, body);
}
export function refillQuote(slide: ISlide, props: QuoteProps, stencil: TemplateSlideInventory): void {
  refill(slide, stencil, `“${props.quote}”`, [props.author, props.role].filter(Boolean).join(", "));
}
export function refillSectionDivider(slide: ISlide, props: SectionDividerProps, stencil: TemplateSlideInventory): void {
  refill(slide, stencil, props.title, props.subtitle);
}
export function refillClosingNextSteps(slide: ISlide, props: ClosingNextStepsProps, stencil: TemplateSlideInventory): void {
  refill(slide, stencil, props.title, props.steps.map((s, i) => `${i + 1}. ${s.title}${s.owner ? ` - ${s.owner}` : ""}`).join("\n"));
}

// convenience for from-scratch (neutral palette default)
export const NEUTRAL = NEUTRAL_PALETTE;
