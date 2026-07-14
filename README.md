# deckproof-mcp

**Bring your own PowerPoint template - get spec-clean slides back.** A [Model Context Protocol](https://modelcontextprotocol.io) server that creates, validates, audits, and repairs `.pptx` files against the real OOXML / ECMA-376 (ISO/IEC 29500) PresentationML spec.

Point it at your company deck; it adds new slides that match your existing masters, layouts, theme, and branding - and every file it hands back is validated against the same structural checklist it uses to inspect other people's files. Fully local. No API keys.

---

## Why this exists

LLM-driven and library-driven PowerPoint generation quietly produces broken files. They open fine in PowerPoint (which silently self-heals on load) and then fail to import into Apple Keynote, Google Slides, or Apache POI - because those readers are strict about the spec.

This isn't hypothetical:

- Anthropic's own `/pptx` skill shipped files that fail to open in Keynote ([anthropics/skills#1167](https://github.com/anthropics/skills/issues/1167)) - root-caused to a stale `sldSz` attribute, a missing `notesMasterIdLst`, and a Windows-only embedded part.
- PptxGenJS passes invalid shape presets straight through, and its `defineSlideMaster` can write malformed `[Content_Types].xml` on multi-slide decks.
- pptx-automizer (the template-cloning engine used here) can leave dangling relationships and unregistered slide IDs.

deckproof-mcp treats the OOXML spec as the source of truth: it generates against it, checks against it, and repairs to it. Creation is **self-certifying** - every generated or repaired file is run back through the validator before it is returned.

## Install

Run it from npm with `npx` (no install needed) and register it with your MCP client:

```json
{
  "mcpServers": {
    "deckproof": {
      "command": "npx",
      "args": ["-y", "deckproof-mcp"]
    }
  }
}
```

> **Running from source** (before the npm package is published, or to hack on it): clone this repo, `npm install && npm run build`, then point your client at `"command": "node", "args": ["/absolute/path/to/deckproof-mcp/dist/index.js"]`.

## Tools

| Tool | What it does |
| --- | --- |
| `pptx_list_layouts` | List the 15 built-in layout archetypes. Pass a `template` to also list that file's slides (with placeholder types) so you can pick stencils. |
| `pptx_create_deck` | Build a deck - from your uploaded template (clone + refill, branding inherited) or from a neutral theme. Self-certifying. |
| `pptx_validate` | Structural pass/fail for any `.pptx` against the 14-rule OOXML checklist. |
| `pptx_audit` | Validation plus advisory metrics: alt-text coverage, orphaned/duplicate media, counts, and per-viewer portability risk. |
| `pptx_repair` | Auto-fix an existing `.pptx`'s structural problems and report what could not be fixed safely. |

Files are returned as a base64 resource block (correct PresentationML MIME type) plus a text summary; structured JSON reports come back in `structuredContent`.

## Bring-your-own-template workflow

1. Call `pptx_list_layouts` with your company `.pptx` as `template`. You get back its slides, each with an `index` and its placeholder types.
2. Call `pptx_create_deck` with the same `template` and a `slides` array. For each slide, choose an `archetype` (the content shape) and a `stencilSlideIndex` (which of your template's slides to clone for its look).
3. Text archetypes (cover, agenda, bullets, quote, ...) refill the cloned slide's placeholders. Rich archetypes (tables, timeline, org chart, matrix, ...) keep the cloned slide's branded chrome and generate their content on top, styled with colors pulled from your template's theme.

New slides are **appended** after your template's existing slides (this preserves spec-correctness), so your template must contain at least one slide to clone. Omit `template` entirely to build a fresh deck on a neutral theme.

### The 15 layout archetypes

cover, agenda, contentBullets, twoColumns, quote, sectionDivider, closingNextSteps, comparisonTable, dataTable, timeline, statsBanner, cardGrid, orgChart, matrixQuadrant, verticalSteps.

## What gets validated (14 rules)

**Structural errors** (gate `valid`): dangling relationships; missing presentation ID-list registrations (`sldIdLst` / `sldMasterIdLst` / `notesMasterIdLst`); content-type completeness; stale/contradictory `sldSz`; slide-layout-master inheritance; duplicate slide IDs; theme color-map breakage; orphaned master/layout chains; duplicate relationship IDs; unregistered slide/layout/master/theme parts.

**Advisory warnings** (surfaced by `pptx_audit`): Windows-only platform parts; orphaned media; missing picture alt-text; duplicate media bloat.

## Deploying remotely (Streamable HTTP)

The same binary also speaks [Streamable HTTP](https://modelcontextprotocol.io/specification/2025-06-18/basic/transports#streamable-http) so it can run as a deployable service. `node dist/index.js` picks the transport at runtime:

- no `PORT` / `MCP_TRANSPORT` → **stdio** (the `npx` default).
- `PORT` set (most PaaS set this automatically), or `MCP_TRANSPORT=http` → **Streamable HTTP** at `POST /mcp`, with a `GET /healthz` check.

The HTTP server is **stateless** (fresh server per request) - safe because every tool call is a pure function of its inputs; the uploaded template travels in the request bytes, not in server state.

```bash
npm run build
npm run start:http      # MCP_TRANSPORT=http PORT=3000 node dist/index.js
# or:
docker build -t deckproof-mcp .
docker run -p 3000:3000 deckproof-mcp
```

### Environment variables (HTTP mode)

| Variable | Default | Purpose |
| --- | --- | --- |
| `PORT` | `3000` | Port to listen on (setting it alone switches to HTTP mode). |
| `MCP_TRANSPORT` | (unset) | Set to `http` to force HTTP mode without `PORT`. |
| `HOST` | `0.0.0.0` | Interface to bind. |
| `MCP_ALLOWED_HOSTS` | (unset) | Comma-separated allowed `Host` header values (DNS-rebinding protection). |

## Security

This server ingests untrusted `.pptx` uploads, so it is hardened accordingly:

- **`path` input is stdio-only.** Reading a file from a local path is allowed only on the local stdio transport. Over HTTP the `path` input is rejected (local-file-read / SSRF protection) - remote callers must pass `base64`.
- **Input limits.** Uploads over 50 MB are rejected; packages that decompress past 300 MB are refused (zip-bomb guard); a single call is capped at 200 slides.
- **XML entity expansion is disabled** (no "billion laughs" amplification on crafted parts).
- **No built-in authentication.** These tools are read-only pure computations with no network or credential access, but if you expose the HTTP server beyond a trusted network, put it behind a gateway that handles auth and set `MCP_ALLOWED_HOSTS`.

## Development

```bash
npm install
npm run build        # tsc -> dist/
npm test             # build + vitest (engine, tools, HTTP tiers)
npm run inspect      # MCP Inspector against the built server
```

Architecture: pure, dependency-free logic in `src/engine/` (OPC package model, one file per validation rule, sanitizer, layout composers) with thin MCP wrappers in `src/tools/`. Tests build their fixtures in code (no checked-in binaries), so they never drift from what the engine emits.

## How it works

- `jszip` + `fast-xml-parser` model the `.pptx` as an OPC package; every rule and fixer operates on that model, never raw bytes.
- `pptxgenjs` builds from-scratch slides; `pptx-automizer` clones template stencils and refills placeholders (targeting them by standard placeholder type + `nameIdx`, robust to duplicate element names).
- The validator runs 14 pure `(package) => violations[]` rules; the sanitizer runs the safe subset as fixers and re-validates.

## License

MIT - see [LICENSE](./LICENSE).
