/**
 * pptx_validate - structural pass/fail against the OOXML validation
 * checklist only, on ANY existing .pptx (from any source, not just files
 * this server created). No repair, no advisory opinions - that's pptx_audit
 * and pptx_repair.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { fileSourceSchema, violationSchema } from "../schemas.js";
import { readFileSource } from "../fileIo.js";
import { validateBuffer } from "../engine/validate.js";
import { formatViolations } from "../format.js";
import { safe, type ServerContext } from "./_util.js";

export function registerPptxValidate(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "pptx_validate",
    {
      title: "Validate a .pptx against the OOXML spec",
      description:
        "Check an existing .pptx file's structural conformance to the OOXML/ECMA-376 PresentationML spec: dangling relationships, missing presentation ID-list registrations (the class of bug that makes files silently fail in strict readers like Apple Keynote), and content-type completeness. Works on any .pptx, regardless of what tool created it. Returns `valid: true` only if there are zero error-severity violations.",
      inputSchema: {
        source: fileSourceSchema,
      },
      outputSchema: {
        valid: z.boolean(),
        violations: z.array(violationSchema),
      },
      annotations: {
        readOnlyHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    safe(async (args) => {
      const bytes = await readFileSource(args.source, ctx);
      const result = await validateBuffer(bytes);

      return {
        content: [
          {
            type: "text",
            text: `${result.valid ? "VALID" : "INVALID"} - ${formatViolations(result.violations)}`,
          },
        ],
        structuredContent: { valid: result.valid, violations: result.violations },
      };
    })
  );
}
