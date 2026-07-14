/**
 * Reads an uploaded template's existing slides so a caller can pick which one
 * to use as a stencil per new slide. Placeholder types come back as the
 * STANDARD OOXML values (title / ctrTitle / subTitle / body / ...), so stencil
 * matching is by spec type, never by template-specific element naming.
 */

import { Automizer } from "pptx-automizer";

export interface TemplateElement {
  /** pptx-automizer element name (used as the modifyElement selector). */
  name: string;
  /** Occurrence index disambiguating duplicate names (e.g. two "Text 0"s). */
  nameIdx: number;
  /** creationId when the template has one (robust selector; PowerPoint sets these). */
  creationId?: string;
  /** Standard OOXML placeholder type, if this element is a placeholder. */
  placeholderType?: string;
}

export interface TemplateSlideInventory {
  /** 0-based index; pass this as `stencilSlideIndex` to pptx_create_deck. */
  index: number;
  layoutName: string;
  elements: TemplateElement[];
}

const TEMPLATE_NAME = "template";

export async function inspectTemplate(bytes: Buffer): Promise<TemplateSlideInventory[]> {
  const automizer = new Automizer({});
  automizer.load(bytes, TEMPLATE_NAME);

  let infos;
  try {
    infos = await automizer.setCreationIds();
  } catch (err) {
    throw new Error(
      `Could not read the uploaded file as a .pptx template: ${err instanceof Error ? err.message : String(err)}`
    );
  }

  const tpl = infos.find((t) => t.name === TEMPLATE_NAME) ?? infos[0];
  if (!tpl) return [];

  return tpl.slides.map((slide, index) => ({
    index,
    layoutName: slide.info?.layoutName ?? "",
    elements: slide.elements.map((el) => {
      let placeholderType: string | undefined;
      try {
        placeholderType = el.getPlaceholderInfo()?.type;
      } catch {
        placeholderType = undefined;
      }
      return {
        name: el.name,
        nameIdx: el.nameIdx ?? 0,
        creationId: el.creationId || undefined,
        placeholderType,
      };
    }),
  }));
}
