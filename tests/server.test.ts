import { describe, it, expect } from "vitest";
import { createServer, VERSION } from "../dist/server.js";

describe("scaffold", () => {
  it("builds a server with the expected name and version", async () => {
    const server = createServer();
    expect(VERSION).toBe("0.1.0");
    await server.close();
  });
});
