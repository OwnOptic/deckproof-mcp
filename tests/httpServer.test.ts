import { describe, it, expect, beforeAll, afterAll } from "vitest";
import type { Server as HttpServer } from "node:http";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { startHttpServer } from "../dist/httpServer.js";

let httpServer: HttpServer;
let baseUrl: string;

beforeAll(async () => {
  httpServer = await startHttpServer({ port: 0, host: "127.0.0.1" });
  const address = httpServer.address();
  if (address === null || typeof address === "string") throw new Error("Expected AddressInfo.");
  baseUrl = `http://127.0.0.1:${address.port}`;
});

afterAll(async () => {
  await new Promise<void>((resolve) => httpServer.close(() => resolve()));
});

describe("Streamable HTTP transport", () => {
  it("answers a health check", async () => {
    const res = await fetch(`${baseUrl}/healthz`);
    expect(res.status).toBe(200);
    const body = await res.json();
    expect(body.server).toBe("deckproof-mcp");
  });

  it("rejects GET and DELETE on /mcp (stateless)", async () => {
    expect((await fetch(`${baseUrl}/mcp`)).status).toBe(405);
    expect((await fetch(`${baseUrl}/mcp`, { method: "DELETE" })).status).toBe(405);
  });

  it("serves a full MCP session: list tools and create a deck", async () => {
    const transport = new StreamableHTTPClientTransport(new URL(`${baseUrl}/mcp`));
    const client = new Client({ name: "http-test", version: "1.0.0" });
    await client.connect(transport);

    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name)).toContain("pptx_create_deck");

    const res: any = await client.callTool({
      name: "pptx_create_deck",
      arguments: { slides: [{ archetype: "cover", title: "Over HTTP" }] },
    });
    expect(res.structuredContent.valid).toBe(true);

    await client.close();
  });

  it("rejects a local path over HTTP (transport is not stdio)", async () => {
    const transport = new StreamableHTTPClientTransport(new URL(`${baseUrl}/mcp`));
    const client = new Client({ name: "http-test-2", version: "1.0.0" });
    await client.connect(transport);
    const res: any = await client.callTool({
      name: "pptx_validate",
      arguments: { source: { path: "C:/whatever.pptx" } },
    });
    expect(res.isError).toBe(true);
    await client.close();
  });

  it("is stateless: independent connections each work standalone", async () => {
    for (let i = 0; i < 2; i++) {
      const transport = new StreamableHTTPClientTransport(new URL(`${baseUrl}/mcp`));
      const client = new Client({ name: `c${i}`, version: "1.0.0" });
      await client.connect(transport);
      const { tools } = await client.listTools();
      expect(tools.length).toBe(5);
      await client.close();
    }
  });
});
