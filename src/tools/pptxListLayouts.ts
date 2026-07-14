/**
 * pptx_list_layouts - discoverability tool. Always returns the static
 * archetype catalog. When an optional `template` is supplied, it ALSO returns
 * that file's existing slide inventory (index, layout name, placeholder types)
 * so the caller can choose which slide to clone as a stencil per new slide in
 * pptx_create_deck.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { listLayouts } from "../engine/layouts/index.js";
import { inspectTemplate } from "../engine/templateInspect.js";
import { fileSourceSchema } from "../schemas.js";
import { readFileSource } from "../fileIo.js";
import { safe, type ServerContext } from "./_util.js";

export function registerPptxListLayouts(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "pptx_list_layouts",
    {
      title: "List layout archetypes (and optionally a template's stencils)",
      description:
        "Returns the catalog of slide layout archetypes pptx_create_deck supports (e.g. cover, agenda, comparison table). If you pass a `template` (an existing .pptx), it also returns that template's existing slides with their layout names and placeholder types, so you can pick a `stencilSlideIndex` for each new slide in template mode.",
      inputSchema: {
        template: fileSourceSchema
          .optional()
          .describe("Optional existing .pptx to inspect for stencil slides."),
      },
      outputSchema: {
        layouts: z.array(z.object({ archetype: z.string(), description: z.string() })),
        templateSlides: z
          .array(
            z.object({
              index: z.number().int(),
              layoutName: z.string(),
              placeholderTypes: z.array(z.string()),
            })
          )
          .optional(),
      },
      annotations: {
        readOnlyHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    safe(async (args) => {
      const layouts = listLayouts();
      const lines = [`${layouts.length} layout archetype(s):`];
      for (const l of layouts) lines.push(`- ${l.archetype}: ${l.description}`);

      let templateSlides:
        | Array<{ index: number; layoutName: string; placeholderTypes: string[] }>
        | undefined;

      if (args.template) {
        const bytes = await readFileSource(args.template, ctx);
        const inventory = await inspectTemplate(bytes);
        templateSlides = inventory.map((s) => ({
          index: s.index,
          layoutName: s.layoutName,
          placeholderTypes: s.elements
            .map((e) => e.placeholderType)
            .filter((t): t is string => Boolean(t)),
        }));
        lines.push("", `Template has ${templateSlides.length} slide(s):`);
        for (const s of templateSlides) {
          lines.push(
            `- [${s.index}] layout "${s.layoutName}", placeholders: ${s.placeholderTypes.join(", ") || "(none)"}`
          );
        }
      }

      return {
        content: [{ type: "text", text: lines.join("\n") }],
        structuredContent: templateSlides ? { layouts, templateSlides } : { layouts },
      };
    })
  );
}
