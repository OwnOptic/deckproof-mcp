/**
 * Thin wrapper around pptx-automizer for the bring-your-own-template path.
 *
 * The uploaded template is loaded twice: as the ROOT (so the output inherits
 * its masters/layouts/theme and keeps the user's own slides) and as a named
 * SOURCE to clone stencil slides from. New slides are APPENDED after the
 * template's existing slides - deliberately NOT using pptx-automizer's
 * `removeExistingSlides`, which corrupts the package (dangling slide
 * relationships + unregistered slide IDs; verified against our own validator).
 * `autoImportSlideMasters` guarantees each cloned slide's master/layout chain
 * travels with it. Output is still run back through our validate/sanitize gate,
 * since pptx-automizer does not guarantee spec-clean output in general.
 */

import { Automizer, ModifyTextHelper } from "pptx-automizer";

export const SOURCE_NAME = "source";

export function createTemplateAutomizer(templateBytes: Buffer): Automizer {
  const automizer = new Automizer({
    autoImportSlideMasters: true,
    cleanup: true,
  });
  automizer.loadRoot(templateBytes);
  automizer.load(templateBytes, SOURCE_NAME);
  return automizer;
}

/** Finalize an Automizer instance and return the resulting .pptx as a Buffer. */
export async function automizerToBuffer(automizer: Automizer): Promise<Buffer> {
  const zip = await automizer.getJSZip();
  return zip.generateAsync({ type: "nodebuffer" });
}

/**
 * Build a robust element selector for modifyElement: prefer a creationId when
 * the template has one (PowerPoint-authored files do), otherwise a
 * name+nameIdx pair, which disambiguates duplicate element names (e.g. two
 * shapes both named "Text 0" in PptxGenJS-generated templates).
 */
export function elementSelector(el: {
  name: string;
  nameIdx: number;
  creationId?: string;
}): string | { name: string; nameIdx: number; creationId?: string } {
  if (el.creationId) return { name: el.name, nameIdx: el.nameIdx, creationId: el.creationId };
  return { name: el.name, nameIdx: el.nameIdx };
}

export { ModifyTextHelper };
