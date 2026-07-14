/**
 * pptx_repair - run the generic sanitize pass over an existing .pptx and
 * return the repaired file plus a log of what changed and any violations that
 * could not be fixed safely. Only makes safe changes (removes things that
 * point nowhere, registers things that exist, normalizes invalid values); it
 * never invents content.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { fileSourceSchema, violationSchema } from "../schemas.js";
import { readFileSource } from "../fileIo.js";
import { sanitizeBuffer } from "../engine/sanitize.js";
import { formatViolations } from "../format.js";
import { safe, type ServerContext } from "./_util.js";

const PPTX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export function registerPptxRepair(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "pptx_repair",
    {
      title: "Repair a .pptx against the OOXML spec",
      description:
        "Auto-fix an existing .pptx's structural problems: backfill missing presentation ID-list entries (the Keynote-import bug), add/correct missing content-type declarations, remove dangling relationships and orphaned/duplicate media, strip Windows-only printer-settings parts, and clamp invalid geometry. Returns the repaired file plus a log of changes made and any violations it could not fix safely (which it reports rather than guessing at).",
      inputSchema: {
        source: fileSourceSchema,
      },
      outputSchema: {
        changesApplied: z.array(z.object({ ruleId: z.string(), description: z.string(), partName: z.string().optional() })),
        remainingViolations: z.array(violationSchema),
        stillValid: z.boolean().describe("True if no error-severity violations remain after repair."),
        byteLength: z.number().int(),
      },
      annotations: {
        readOnlyHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    safe(async (args) => {
      const bytes = await readFileSource(args.source, ctx);
      const { bytes: repaired, changes, remainingViolations } = await sanitizeBuffer(bytes);
      const stillValid = !remainingViolations.some((v) => v.severity === "error");
      const base64 = Buffer.from(repaired).toString("base64");

      const text = [
        `Applied ${changes.length} fix(es).`,
        ...changes.map((c) => `  + ${c.description}`),
        `After repair: ${stillValid ? "VALID" : "STILL INVALID"} - ${formatViolations(remainingViolations)}`,
      ].join("\n");

      return {
        content: [
          { type: "text", text },
          {
            type: "resource",
            resource: { uri: "deckproof://repaired-deck.pptx", mimeType: PPTX_MIME_TYPE, blob: base64 },
          },
        ],
        structuredContent: {
          changesApplied: changes,
          remainingViolations,
          stillValid,
          byteLength: repaired.byteLength,
        },
      };
    })
  );
}
