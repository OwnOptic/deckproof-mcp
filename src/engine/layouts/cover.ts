/**
 * Cover slide archetype: eyebrow + title + subtitle + optional context line.
 * Works both as a from-scratch PptxGenJS composition and (once template mode
 * lands) as content poured into a cloned stencil's existing placeholders.
 */

import { z } from "zod";
import type { ISlide } from "pptx-automizer";
import type { PptxSlide } from "../pptxgenTypes.js";
import type { TemplateSlideInventory } from "../templateInspect.js";
import { ModifyTextHelper, elementSelector } from "../automizer.js";

export const coverPropsSchema = z.object({
  archetype: z.literal("cover"),
  title: z.string().min(1).describe("Main deck or section title."),
  eyebrow: z.string().optional().describe("Small label above the title, e.g. a category or date."),
  subtitle: z.string().optional().describe("Supporting line under the title."),
  context: z.string().optional().describe("Small footer-style context line, e.g. author or audience."),
});

export type CoverProps = z.infer<typeof coverPropsSchema>;

/** Compose a cover slide from scratch (no template). Coordinates assume a 13.33x7.5in widescreen slide. */
export function composeCoverFromScratch(slide: PptxSlide, props: CoverProps): void {
  if (props.eyebrow) {
    slide.addText(props.eyebrow, {
      x: 0.6, y: 1.6, w: 12, h: 0.5,
      fontSize: 14, color: "6B7280", bold: true, charSpacing: 2,
    });
  }
  slide.addText(props.title, {
    x: 0.6, y: 2.1, w: 12, h: 1.6,
    fontSize: 40, bold: true, color: "111827",
  });
  if (props.subtitle) {
    slide.addText(props.subtitle, {
      x: 0.6, y: 3.7, w: 11, h: 0.8,
      fontSize: 18, color: "374151",
    });
  }
  if (props.context) {
    slide.addText(props.context, {
      x: 0.6, y: 6.9, w: 12, h: 0.4,
      fontSize: 11, color: "9CA3AF",
    });
  }
}

/**
 * Template mode: refill a cloned cover stencil's placeholders. Title text goes
 * to the title/ctrTitle placeholder; subtitle to the subTitle placeholder (or
 * the body placeholder if there's no subtitle placeholder). Branding, fonts,
 * and colors are inherited from the cloned slide - we only replace text.
 */
export function refillCoverStencil(
  slide: ISlide,
  props: CoverProps,
  stencil: TemplateSlideInventory
): void {
  const titleEl = stencil.elements.find(
    (e) => e.placeholderType === "title" || e.placeholderType === "ctrTitle"
  );
  if (titleEl) {
    slide.modifyElement(elementSelector(titleEl), ModifyTextHelper.setText(props.title));
  }

  const subtitleText = props.subtitle ?? props.eyebrow;
  if (subtitleText) {
    const subEl = stencil.elements.find(
      (e) => e.placeholderType === "subTitle" || e.placeholderType === "body"
    );
    if (subEl) {
      slide.modifyElement(elementSelector(subEl), ModifyTextHelper.setText(subtitleText));
    }
  }
}
