import { describe, it, expect, beforeAll, afterAll } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../dist/index.js";
import { buildSyntheticTemplate, ALL_FIFTEEN_SLIDES } from "./fixtures.js";

// Two servers: one HTTP-like (no local paths), one stdio-like (paths allowed).
let httpClient: Client;
let stdioClient: Client;
let templateBase64: string;

async function connect(allowLocalPaths: boolean): Promise<Client> {
  const [ct, st] = InMemoryTransport.createLinkedPair();
  const server = createServer({ allowLocalPaths });
  await server.connect(st);
  const client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(ct);
  return client;
}

function resourceBlob(res: any): Buffer {
  const block = res.content.find((c: any) => c.type === "resource");
  return Buffer.from(block.resource.blob, "base64");
}

beforeAll(async () => {
  httpClient = await connect(false);
  stdioClient = await connect(true);
  templateBase64 = (await buildSyntheticTemplate()).toString("base64");
});

afterAll(async () => {
  await httpClient.close();
  await stdioClient.close();
});

describe("tool registration", () => {
  it("exposes all five tools", async () => {
    const { tools } = await httpClient.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual(
      ["pptx_audit", "pptx_create_deck", "pptx_list_layouts", "pptx_repair", "pptx_validate"].sort()
    );
  });
});

describe("pptx_list_layouts", () => {
  it("returns all 15 archetypes", async () => {
    const res: any = await httpClient.callTool({ name: "pptx_list_layouts", arguments: {} });
    expect(res.structuredContent.layouts).toHaveLength(15);
  });

  it("returns a template's stencil inventory when a template is supplied", async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_list_layouts",
      arguments: { template: { base64: templateBase64 } },
    });
    expect(res.structuredContent.templateSlides.length).toBeGreaterThan(0);
    expect(res.structuredContent.templateSlides[0].placeholderTypes).toContain("title");
  });
});

describe("pptx_create_deck - from scratch", () => {
  it("builds a valid deck using all 15 archetypes", async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_create_deck",
      arguments: { title: "All 15", slides: ALL_FIFTEEN_SLIDES },
    });
    expect(res.structuredContent.mode).toBe("from-scratch");
    expect(res.structuredContent.valid).toBe(true);
    expect(res.structuredContent.slidesAdded).toBe(15);
    const buf = resourceBlob(res);
    expect(buf[0]).toBe(0x50); // 'P' - zip magic
    expect(buf[1]).toBe(0x4b); // 'K'
  });

  it("rejects a deck over the slide cap", async () => {
    const many = Array.from({ length: 201 }, () => ({ archetype: "quote", quote: "x" }));
    const res: any = await httpClient.callTool({ name: "pptx_create_deck", arguments: { slides: many } });
    expect(res.isError).toBe(true);
  });
});

describe("pptx_create_deck - template mode", () => {
  it("appends valid, refilled slides cloned from the template", async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_create_deck",
      arguments: {
        template: { base64: templateBase64 },
        slides: [
          { archetype: "contentBullets", stencilSlideIndex: 0, title: "Refilled", bullets: ["one", "two"] },
          { archetype: "statsBanner", stencilSlideIndex: 0, stats: [{ value: "42", label: "answer" }] },
        ],
      },
    });
    expect(res.structuredContent.mode).toBe("template");
    expect(res.structuredContent.valid).toBe(true);
    expect(res.structuredContent.slidesAdded).toBe(2);
  });

  it("errors when a template slide is missing stencilSlideIndex", async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_create_deck",
      arguments: { template: { base64: templateBase64 }, slides: [{ archetype: "quote", quote: "x" }] },
    });
    expect(res.isError).toBe(true);
  });
});

describe("pptx_validate / pptx_audit", () => {
  let validDeckB64: string;
  beforeAll(async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_create_deck",
      arguments: { slides: [{ archetype: "cover", title: "V" }] },
    });
    validDeckB64 = resourceBlob(res).toString("base64");
  });

  it("validates a clean deck", async () => {
    const res: any = await httpClient.callTool({ name: "pptx_validate", arguments: { source: { base64: validDeckB64 } } });
    expect(res.structuredContent.valid).toBe(true);
  });

  it("audits a deck with portability metrics", async () => {
    const res: any = await httpClient.callTool({ name: "pptx_audit", arguments: { source: { base64: validDeckB64 } } });
    const m = res.structuredContent.metrics;
    expect(m.slideCount).toBeGreaterThanOrEqual(1);
    expect(m.altTextCoveragePct).toBe(100); // no pictures -> 100%
    expect(["low", "medium", "high"]).toContain(m.portabilityRisk.keynote);
  });

  it("returns a clean error on malformed input (no crash)", async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_validate",
      arguments: { source: { base64: Buffer.from("not a pptx").toString("base64") } },
    });
    expect(res.isError).toBe(true);
  });
});

describe("pptx_repair", () => {
  it("repairs a deck broken by a stripped content-type override", async () => {
    // Build valid, corrupt via the low-level engine, hand the bytes to the tool.
    const { createFromScratch } = await import("../dist/engine/create.js");
    const { loadPackage, savePackage } = await import("../dist/engine/opcPackage.js");
    const good = await createFromScratch({ slides: [{ archetype: "cover", title: "R" }] });
    const pkg = await loadPackage(good);
    pkg.contentTypes = pkg.contentTypes.filter(
      (e: any) => !(e.kind === "override" && e.partName === "/ppt/slides/slide1.xml")
    );
    const broken = await savePackage(pkg);

    const res: any = await httpClient.callTool({
      name: "pptx_repair",
      arguments: { source: { base64: Buffer.from(broken).toString("base64") } },
    });
    expect(res.structuredContent.changesApplied.length).toBeGreaterThan(0);
    expect(res.structuredContent.stillValid).toBe(true);
  });
});

describe("path security", () => {
  it("rejects a local path over the HTTP-like (allowLocalPaths=false) server", async () => {
    const res: any = await httpClient.callTool({
      name: "pptx_validate",
      arguments: { source: { path: "C:/nonexistent/whatever.pptx" } },
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/stdio/i);
  });

  it("attempts the read over the stdio-like (allowLocalPaths=true) server", async () => {
    // Path is allowed, so it gets past the guard and fails on the missing file
    // instead - a different error, proving the guard let it through.
    const res: any = await stdioClient.callTool({
      name: "pptx_validate",
      arguments: { source: { path: "C:/nonexistent/whatever.pptx" } },
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).not.toMatch(/only available on the local stdio/i);
  });
});
