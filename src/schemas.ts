/**
 * Shared Zod schemas for tool inputs and outputs.
 *
 * Input shapes are exported as raw shapes (plain objects of Zod types) so
 * they can be passed straight to `server.registerTool`. Output shapes are
 * likewise raw shapes used as `outputSchema`, which the SDK validates
 * `structuredContent` against.
 */

import { z } from "zod";

/** A .pptx supplied either as a local filesystem path or base64-encoded bytes. */
export const fileSourceSchema = z
  .object({
    path: z
      .string()
      .optional()
      .describe("Local filesystem path to a .pptx file (stdio/local use only)."),
    base64: z
      .string()
      .optional()
      .describe("Base64-encoded .pptx file bytes (works over any transport, including remote HTTP)."),
  })
  .refine((o) => Boolean(o.path) !== Boolean(o.base64), {
    message: "Provide exactly one of `path` or `base64`.",
  });

export type FileSource = z.infer<typeof fileSourceSchema>;

export const violationSchema = z.object({
  ruleId: z.string().describe("Stable identifier of the rule that fired, e.g. 'danglingRelationships'."),
  severity: z.enum(["error", "warning"]),
  message: z.string(),
  partName: z.string().optional().describe("OPC part name the violation was found in, if applicable."),
});

export type Violation = z.infer<typeof violationSchema>;
