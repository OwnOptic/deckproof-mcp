/**
 * Shared tool-layer helpers.
 *
 * `ServerContext` carries transport-derived capabilities (currently just
 * whether local filesystem paths are allowed) from the server bootstrap down
 * to individual tools.
 *
 * `safe()` wraps a tool handler so any thrown error becomes a clean MCP
 * `isError` result with an actionable message, rather than an unhandled
 * rejection that would surface as an opaque protocol error. Essential when
 * processing untrusted uploads that may be malformed in countless ways.
 */

export interface ServerContext {
  allowLocalPaths: boolean;
}

/**
 * Wrap a tool handler so thrown errors become a clean `isError` result.
 * Preserves the handler's exact return type R (so the SDK's tool-result type
 * inference still holds); the error branch is a valid CallToolResult and is
 * cast to R at that single, justified point.
 */
export function safe<Args, R>(
  handler: (args: Args) => Promise<R>
): (args: Args) => Promise<R> {
  return async (args: Args): Promise<R> => {
    try {
      return await handler(args);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        isError: true,
        content: [{ type: "text", text: `Error: ${message}` }],
      } as unknown as R;
    }
  };
}
