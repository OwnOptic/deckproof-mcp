/**
 * pptx_create_deck - stateless, single-call deck creation. Takes a declarative
 * list of slides and returns a complete .pptx in one round trip (no
 * server-side session, so this stays deployment-friendly - the "session" is
 * the request/response bytes, never anything the server remembers).
 *
 * Two modes:
 *  - Bring-your-own-template: pass `template` (an existing company .pptx) and a
 *    `stencilSlideIndex` per slide; each new slide is cloned from that stencil
 *    (inheriting masters/layout/theme/branding) and refilled.
 *  - From-scratch: omit `template`; slides are built on a neutral theme.
 *
 * Self-certifying: always runs the same validator used by pptx_validate on its
 * own output before returning, so `valid: true` is a guarantee, not an
 * assumption (both PptxGenJS and pptx-automizer have documented silent-failure
 * bugs, which is exactly why the gate exists).
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createFromScratch, createFromTemplate, type TemplateSlideSpec } from "../engine/create.js";
import { slideInputSchema } from "../engine/layouts/index.js";
import { sanitizeBuffer } from "../engine/sanitize.js";
import { fileSourceSchema, violationSchema } from "../schemas.js";
import { readFileSource } from "../fileIo.js";
import { formatViolations } from "../format.js";
import { safe, type ServerContext } from "./_util.js";

const PPTX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_SLIDES = 200;

export function registerPptxCreateDeck(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "pptx_create_deck",
    {
      title: "Create a PowerPoint deck",
      description:
        "Build a .pptx from a declarative list of slides (see pptx_list_layouts for the archetype catalog). Bring-your-own-template mode: pass `template` (your existing company .pptx) and a `stencilSlideIndex` per slide - each new slide is cloned from that template slide, inheriting its masters, layout, theme and branding, then refilled with your content. New slides are APPENDED after the template's existing slides (this preserves spec-correctness; the template must contain at least one slide to clone). Omit `template` to build a fresh deck from a neutral default theme. Every call is self-certifying: output is validated against the same OOXML checklist pptx_validate uses before it's returned. Returns the file as a base64 resource plus a summary.",
      inputSchema: {
        title: z.string().optional().describe("Presentation title (metadata, not a slide)."),
        template: fileSourceSchema
          .optional()
          .describe("Optional existing .pptx to use as the brand template. When given, every slide needs a stencilSlideIndex."),
        slides: z.array(slideInputSchema).min(1).describe("The slides to build, in order."),
      },
      outputSchema: {
        mode: z.enum(["template", "from-scratch"]),
        slidesAdded: z.number().int().describe("Number of new slides created (appended, in template mode)."),
        byteLength: z.number().int(),
        valid: z.boolean(),
        violations: z.array(violationSchema),
      },
      annotations: {
        readOnlyHint: false,
        idempotentHint: false,
        openWorldHint: false,
      },
    },
    safe(async (args) => {
      if (args.slides.length > MAX_SLIDES) {
        throw new Error(`Too many slides (${args.slides.length}); the limit is ${MAX_SLIDES} per call.`);
      }
      let buffer: Uint8Array;
      let mode: "template" | "from-scratch";

      if (args.template) {
        mode = "template";
        const templateBytes = await readFileSource(args.template, ctx);
        const missing = args.slides.findIndex(
          (s) => (s as { stencilSlideIndex?: number }).stencilSlideIndex === undefined
        );
        if (missing !== -1) {
          throw new Error(
            `Template mode requires a stencilSlideIndex on every slide; slide ${missing + 1} is missing one. Call pptx_list_layouts with the same template to see available stencil slides.`
          );
        }
        buffer = await createFromTemplate({
          templateBytes,
          title: args.title,
          slides: args.slides as unknown as TemplateSlideSpec[],
        });
      } else {
        mode = "from-scratch";
        buffer = await createFromScratch({ title: args.title, slides: args.slides });
      }

      // Mandatory self-certifying gate: sanitize (repair anything the
      // generation libraries got wrong) then report residual violations.
      const { bytes: finalBytes, changes, remainingViolations } = await sanitizeBuffer(buffer);
      const valid = !remainingViolations.some((v) => v.severity === "error");
      const violations = remainingViolations;
      const base64 = Buffer.from(finalBytes).toString("base64");

      const summary =
        mode === "template"
          ? `Appended ${args.slides.length} slide(s) to the template`
          : `Created a ${args.slides.length}-slide deck`;
      const gate = changes.length > 0 ? ` (${changes.length} auto-fix(es) applied)` : "";
      const text = [
        `${summary} (${mode} mode, ${finalBytes.byteLength.toLocaleString("en-US")} bytes).`,
        `${valid ? "VALID" : "INVALID"}${gate} - ${formatViolations(violations)}`,
      ].join("\n");

      return {
        content: [
          { type: "text", text },
          {
            type: "resource",
            resource: {
              uri: "deckproof://generated-deck.pptx",
              mimeType: PPTX_MIME_TYPE,
              blob: base64,
            },
          },
        ],
        structuredContent: {
          mode,
          slidesAdded: args.slides.length,
          byteLength: finalBytes.byteLength,
          valid,
          violations,
        },
      };
    })
  );
}
