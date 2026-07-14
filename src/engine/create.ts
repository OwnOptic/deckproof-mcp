/**
 * Deck creation, two modes:
 *
 * - From-scratch (no template): PptxGenJS with one shared `defineSlideMaster()`
 *   for the whole deck - guards a documented PptxGenJS bug where per-slide
 *   `defineSlideMaster()` writes one <Override> per SLIDE into
 *   [Content_Types].xml instead of per master (rule 5).
 * - From a template: pptx-automizer clones a chosen stencil slide (inheriting
 *   its master/layout/theme/branding) and refills its placeholders, or - for
 *   graphically-rich archetypes - generates content on top styled with the
 *   template's extracted theme colors.
 *
 * Both modes return a raw Buffer; the mandatory sanitize+validate gate is
 * applied by the caller (the create tool), not here.
 */

import PptxGenJSImport from "pptxgenjs";
import { layoutRegistry, type SlideSpec } from "./layouts/index.js";
import type { PptxGenJSConstructor, PptxSlide } from "./pptxgenTypes.js";
import { createTemplateAutomizer, SOURCE_NAME, automizerToBuffer } from "./automizer.js";
import { inspectTemplate } from "./templateInspect.js";
import { extractTheme } from "./themeExtract.js";
import { loadPackage } from "./opcPackage.js";

const PptxGenJS = PptxGenJSImport as unknown as PptxGenJSConstructor;
const MASTER_NAME = "DECKPROOF_MASTER";

export interface CreateFromScratchOptions {
  title?: string;
  slides: SlideSpec[];
}

export interface TemplateSlideSpec {
  archetype: string;
  stencilSlideIndex: number;
  [key: string]: unknown;
}

export interface CreateFromTemplateOptions {
  templateBytes: Buffer;
  title?: string;
  slides: TemplateSlideSpec[];
}

export async function createFromScratch(
  options: CreateFromScratchOptions
): Promise<Uint8Array> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  if (options.title) pptx.title = options.title;

  pptx.defineSlideMaster({
    title: MASTER_NAME,
    background: { color: "FFFFFF" },
  });

  for (const spec of options.slides) {
    const def = layoutRegistry[spec.archetype];
    if (!def) {
      throw new Error(`Unknown layout archetype: ${String((spec as { archetype?: unknown }).archetype)}`);
    }
    const slide = pptx.addSlide({ masterName: MASTER_NAME });
    def.composeFromScratch(slide, spec);
  }

  const output = await pptx.write({ outputType: "nodebuffer" });
  return output as Uint8Array;
}

export async function createFromTemplate(
  options: CreateFromTemplateOptions
): Promise<Uint8Array> {
  const inventory = await inspectTemplate(options.templateBytes);
  if (inventory.length === 0) {
    throw new Error("The uploaded template contains no slides to use as stencils.");
  }
  const theme = extractTheme(await loadPackage(options.templateBytes));

  const automizer = createTemplateAutomizer(options.templateBytes);

  for (const spec of options.slides) {
    const def = layoutRegistry[spec.archetype];
    if (!def) {
      throw new Error(`Unknown layout archetype: ${String(spec.archetype)}`);
    }
    const stencil = inventory[spec.stencilSlideIndex];
    if (!stencil) {
      throw new Error(
        `stencilSlideIndex ${spec.stencilSlideIndex} is out of range; the template has ${inventory.length} slide(s) (valid indexes 0..${inventory.length - 1}).`
      );
    }

    automizer.addSlide(SOURCE_NAME, spec.stencilSlideIndex + 1, (slide) => {
      if (def.refillStencil) {
        def.refillStencil(slide, spec, stencil, theme);
      } else if (def.generateOnStencil) {
        const generate = def.generateOnStencil;
        slide.generate((genSlide: unknown) => {
          generate(genSlide as PptxSlide, spec, theme);
        });
      }
    });
  }

  return automizerToBuffer(automizer);
}
