/**
 * pptx_audit - validation plus advisory quality metrics for an existing
 * .pptx: alt-text coverage, orphaned/duplicate media, slide/media counts,
 * file size, and a per-viewer portability risk estimate (PowerPoint / Keynote
 * / LibreOffice / Google Slides). Read-only and opinion-free beyond the spec.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { fileSourceSchema, violationSchema } from "../schemas.js";
import { readFileSource } from "../fileIo.js";
import { auditBuffer } from "../engine/audit.js";
import { formatViolations } from "../format.js";
import { safe, type ServerContext } from "./_util.js";

const riskEnum = z.enum(["low", "medium", "high"]);

export function registerPptxAudit(server: McpServer, ctx: ServerContext): void {
  server.registerTool(
    "pptx_audit",
    {
      title: "Audit a .pptx (validation + quality metrics)",
      description:
        "Full report on an existing .pptx: all structural validation violations, plus advisory metrics - alt-text coverage %, orphaned and duplicate media, slide/media counts, file size, and an estimated portability risk for PowerPoint, Apple Keynote, LibreOffice Impress, and Google Slides (strict readers reject bugs that lenient ones silently self-heal).",
      inputSchema: {
        source: fileSourceSchema,
      },
      outputSchema: {
        valid: z.boolean(),
        violations: z.array(violationSchema),
        metrics: z.object({
          slideCount: z.number().int(),
          mediaCount: z.number().int(),
          byteLength: z.number().int(),
          pictureCount: z.number().int(),
          altTextCoveragePct: z.number(),
          orphanedMedia: z.array(z.string()),
          duplicateMediaGroups: z.number().int(),
          portabilityRisk: z.object({
            powerpoint: riskEnum,
            keynote: riskEnum,
            libreoffice: riskEnum,
            googleSlides: riskEnum,
          }),
        }),
      },
      annotations: {
        readOnlyHint: true,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    safe(async (args) => {
      const bytes = await readFileSource(args.source, ctx);
      const report = await auditBuffer(bytes);
      const m = report.metrics;
      const risk = m.portabilityRisk;

      const text = [
        `${report.valid ? "VALID" : "INVALID"} - ${report.violations.length} finding(s).`,
        `Slides: ${m.slideCount}  Media: ${m.mediaCount}  Pictures: ${m.pictureCount}  Size: ${(m.byteLength / 1024).toFixed(0)} KB`,
        `Alt-text coverage: ${m.altTextCoveragePct}%  Orphaned media: ${m.orphanedMedia.length}  Duplicate media groups: ${m.duplicateMediaGroups}`,
        `Portability risk - PowerPoint: ${risk.powerpoint}, Keynote: ${risk.keynote}, LibreOffice: ${risk.libreoffice}, Google Slides: ${risk.googleSlides}`,
        "",
        formatViolations(report.violations),
      ].join("\n");

      return {
        content: [{ type: "text", text }],
        structuredContent: {
          valid: report.valid,
          violations: report.violations,
          metrics: m,
        },
      };
    })
  );
}
