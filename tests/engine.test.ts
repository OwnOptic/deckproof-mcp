import { describe, it, expect, beforeAll } from "vitest";
import JSZip from "jszip";
import {
  loadPackage,
  savePackage,
  relsPathFor,
  sourcePartForRelsPath,
  parseRelationshipsXml,
  buildRelationshipsXml,
  parseContentTypesXml,
  buildContentTypesXml,
  getPartText,
  setPartText,
  deletePart,
} from "../dist/engine/opcPackage.js";
import { createFromScratch, createFromTemplate } from "../dist/engine/create.js";
import { checkDanglingRelationships } from "../dist/engine/rules/danglingRelationships.js";
import { checkPresentationIdLists } from "../dist/engine/rules/presentationIdLists.js";
import { checkContentTypesCompleteness } from "../dist/engine/rules/contentTypesCompleteness.js";
import { checkStaleAttributes } from "../dist/engine/rules/staleAttributes.js";
import { checkDuplicateSlideIds } from "../dist/engine/rules/duplicateSlideIds.js";
import { checkPlatformSpecificParts } from "../dist/engine/rules/platformSpecificParts.js";
import { runAllRules } from "../dist/engine/rules/index.js";
import { validateBuffer } from "../dist/engine/validate.js";
import { sanitizeBuffer } from "../dist/engine/sanitize.js";
import type { OpcPackage } from "../dist/engine/types.js";
import { buildSyntheticTemplate } from "./fixtures.js";

async function buildTrivialZip(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>`
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst/></p:presentation>`
  );
  return zip.generateAsync({ type: "uint8array" });
}

describe("relsPathFor / sourcePartForRelsPath", () => {
  it("maps the root part to _rels/.rels", () => {
    expect(relsPathFor("")).toBe("_rels/.rels");
    expect(sourcePartForRelsPath("_rels/.rels")).toBe("");
  });

  it("maps a nested part to its sibling _rels file", () => {
    expect(relsPathFor("ppt/presentation.xml")).toBe("ppt/_rels/presentation.xml.rels");
    expect(sourcePartForRelsPath("ppt/_rels/presentation.xml.rels")).toBe("ppt/presentation.xml");
  });

  it("round-trips slide paths", () => {
    const relsPath = relsPathFor("ppt/slides/slide1.xml");
    expect(relsPath).toBe("ppt/slides/_rels/slide1.xml.rels");
    expect(sourcePartForRelsPath(relsPath)).toBe("ppt/slides/slide1.xml");
  });
});

describe("relationships XML parse/build", () => {
  it("round-trips a relationship list", () => {
    const rels = [
      { id: "rId1", type: "http://example.com/rel", target: "slides/slide1.xml" },
      { id: "rId2", type: "http://example.com/hyperlink", target: "https://example.com", targetMode: "External" as const },
    ];
    const xml = buildRelationshipsXml(rels);
    const parsed = parseRelationshipsXml(xml);
    expect(parsed).toEqual(rels);
  });

  it("parses real OOXML relationship XML", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>`;
    const parsed = parseRelationshipsXml(xml);
    expect(parsed).toEqual([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide", target: "slides/slide1.xml" },
    ]);
  });
});

describe("content-types XML parse/build", () => {
  it("round-trips defaults and overrides", () => {
    const entries = [
      { kind: "default" as const, extension: "xml", contentType: "application/xml" },
      { kind: "override" as const, partName: "/ppt/presentation.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" },
    ];
    const xml = buildContentTypesXml(entries);
    const parsed = parseContentTypesXml(xml);
    expect(parsed).toEqual(entries);
  });
});

describe("loadPackage / savePackage", () => {
  it("round-trips a trivial zip's parts, rels, and content-types unchanged", async () => {
    const bytes = await buildTrivialZip();
    const pkg = await loadPackage(bytes);

    expect(pkg.parts.has("ppt/presentation.xml")).toBe(true);
    expect(pkg.contentTypes).toEqual([
      { kind: "default", extension: "xml", contentType: "application/xml" },
      { kind: "default", extension: "rels", contentType: "application/vnd.openxmlformats-package.relationships+xml" },
      { kind: "override", partName: "/ppt/presentation.xml", contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" },
    ]);
    expect(pkg.relsByPart.get("")).toEqual([
      { id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "ppt/presentation.xml" },
    ]);

    const savedBytes = await savePackage(pkg);
    const reloaded = await loadPackage(savedBytes);

    expect(reloaded.contentTypes).toEqual(pkg.contentTypes);
    expect(reloaded.relsByPart.get("")).toEqual(pkg.relsByPart.get(""));
    expect(getPartText(reloaded, "ppt/presentation.xml")).toBe(
      getPartText(pkg, "ppt/presentation.xml")
    );
  });

  it("setPartText mutations survive a save/reload cycle", async () => {
    const bytes = await buildTrivialZip();
    const pkg = await loadPackage(bytes);
    setPartText(pkg, "ppt/presentation.xml", "<p:presentation>mutated</p:presentation>");

    const saved = await savePackage(pkg);
    const reloaded = await loadPackage(saved);
    expect(getPartText(reloaded, "ppt/presentation.xml")).toBe(
      "<p:presentation>mutated</p:presentation>"
    );
  });
});

describe("validation rules (1-3)", () => {
  let validPkg: OpcPackage;

  beforeAll(async () => {
    const buffer = await createFromScratch({
      slides: [{ archetype: "cover", title: "Rule Test Deck" }],
    });
    validPkg = await loadPackage(buffer);
  });

  it("a clean, freshly-created deck has zero violations across all rules", () => {
    expect(runAllRules(validPkg)).toEqual([]);
  });

  describe("danglingRelationships", () => {
    it("does not fire on a clean deck", () => {
      expect(checkDanglingRelationships(validPkg)).toEqual([]);
    });

    it("fires when a relationship's target part is deleted", async () => {
      const bytes = await savePackage(validPkg);
      const pkg = await loadPackage(bytes);
      const slideRel = pkg.relsByPart.get("ppt/presentation.xml")?.find((r) => r.target.includes("slides/slide1.xml"));
      expect(slideRel).toBeTruthy();
      deletePart(pkg, "ppt/slides/slide1.xml");

      const violations = checkDanglingRelationships(pkg);
      expect(violations.length).toBeGreaterThan(0);
      expect(violations[0].ruleId).toBe("danglingRelationships");
      expect(violations[0].severity).toBe("error");
    });
  });

  describe("presentationIdLists", () => {
    it("does not fire on a clean deck", () => {
      expect(checkPresentationIdLists(validPkg)).toEqual([]);
    });

    it("fires when a slide relationship is not registered in sldIdLst", async () => {
      const bytes = await savePackage(validPkg);
      const pkg = await loadPackage(bytes);
      const text = getPartText(pkg, "ppt/presentation.xml")!;
      // Strip the <p:sldIdLst>...</p:sldIdLst> block entirely, simulating the
      // exact anthropics/skills#1167 class of bug (relationship exists, but
      // presentation.xml never lists it).
      const stripped = text.replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/s, "<p:sldIdLst/>");
      expect(stripped).not.toBe(text);
      setPartText(pkg, "ppt/presentation.xml", stripped);

      const violations = checkPresentationIdLists(pkg);
      expect(violations.length).toBeGreaterThan(0);
      expect(violations[0].ruleId).toBe("presentationIdLists");
      expect(violations[0].severity).toBe("error");
    });
  });

  describe("contentTypesCompleteness", () => {
    it("does not fire on a clean deck", () => {
      expect(checkContentTypesCompleteness(validPkg)).toEqual([]);
    });

    it("fires when a part's content-type declaration is removed", async () => {
      const bytes = await savePackage(validPkg);
      const pkg = await loadPackage(bytes);
      // presentation.xml is covered by BOTH its own <Override> and the generic
      // <Default Extension="xml">; removing only the Override still leaves it
      // covered by the Default, so remove both to make it genuinely uncovered.
      pkg.contentTypes = pkg.contentTypes.filter(
        (e) =>
          !(e.kind === "override" && e.partName === "/ppt/presentation.xml") &&
          !(e.kind === "default" && e.extension === "xml")
      );

      const violations = checkContentTypesCompleteness(pkg);
      expect(violations.length).toBeGreaterThan(0);
      expect(
        violations.some(
          (v) => v.partName === "ppt/presentation.xml" && v.ruleId === "contentTypesCompleteness"
        )
      ).toBe(true);
      expect(violations.every((v) => v.severity === "error")).toBe(true);
    });
  });
});

describe("validation rules (4-14, representative)", () => {
  let validPkg: OpcPackage;
  beforeAll(async () => {
    validPkg = await loadPackage(await createFromScratch({ slides: [{ archetype: "cover", title: "T" }] }));
  });

  it("staleAttributes fires on a contradictory sldSz type", async () => {
    const bytes = await savePackage(validPkg);
    const pkg = await loadPackage(bytes);
    const pres = getPartText(pkg, "ppt/presentation.xml")!;
    // Force a 4x3 type onto 16:9 dimensions.
    const mutated = pres.replace(/<p:sldSz([^>]*?)\/>/, '<p:sldSz cx="9144000" cy="5143500" type="screen4x3"/>');
    setPartText(pkg, "ppt/presentation.xml", mutated);
    const v = checkStaleAttributes(pkg);
    expect(v.length).toBeGreaterThan(0);
    expect(v[0].ruleId).toBe("staleAttributes");
  });

  it("duplicateSlideIds fires on colliding slide ids", async () => {
    const bytes = await savePackage(validPkg);
    const pkg = await loadPackage(bytes);
    const pres = getPartText(pkg, "ppt/presentation.xml")!;
    // Duplicate the single sldId entry.
    const mutated = pres.replace(/(<p:sldId [^>]*\/>)/, "$1$1");
    setPartText(pkg, "ppt/presentation.xml", mutated);
    const v = checkDuplicateSlideIds(pkg);
    expect(v.some((x) => x.ruleId === "duplicateSlideIds")).toBe(true);
  });

  it("platformSpecificParts fires on an injected printerSettings part", async () => {
    const bytes = await savePackage(validPkg);
    const pkg = await loadPackage(bytes);
    pkg.parts.set("ppt/printerSettings/printerSettings1.bin", new Uint8Array([1, 2, 3]));
    const v = checkPlatformSpecificParts(pkg);
    expect(v[0].ruleId).toBe("platformSpecificParts");
    expect(v[0].severity).toBe("warning");
  });
});

describe("sanitize", () => {
  it("repairs a deck broken multiple ways back to valid", async () => {
    const good = await createFromScratch({ slides: [{ archetype: "cover", title: "S" }] });
    const pkg = await loadPackage(good);
    // Strip sldIdLst entries + a slide Override + inject a platform part.
    const pres = getPartText(pkg, "ppt/presentation.xml")!;
    setPartText(pkg, "ppt/presentation.xml", pres.replace(/<p:sldIdLst>.*?<\/p:sldIdLst>/s, "<p:sldIdLst/>"));
    pkg.contentTypes = pkg.contentTypes.filter((e) => !(e.kind === "override" && e.partName === "/ppt/slides/slide1.xml"));
    pkg.parts.set("ppt/printerSettings/printerSettings1.bin", new Uint8Array([0]));
    const broken = await savePackage(pkg);

    expect((await validateBuffer(broken)).valid).toBe(false);
    const { bytes, changes } = await sanitizeBuffer(broken);
    expect(changes.length).toBeGreaterThanOrEqual(2);
    expect((await validateBuffer(bytes)).valid).toBe(true);
  });
});

describe("template mode (clone + refill)", () => {
  it("clones a stencil, refills it, and produces a valid deck", async () => {
    const template = await buildSyntheticTemplate();
    const out = await createFromTemplate({
      templateBytes: template,
      slides: [{ archetype: "cover", stencilSlideIndex: 0, title: "Refilled Title", subtitle: "Refilled sub" }],
    });
    const result = await validateBuffer(out);
    expect(result.valid).toBe(true);
    // The appended (last) slide should carry the refilled title.
    const pkg = await loadPackage(out);
    const slideParts = [...pkg.parts.keys()].filter((p) => /ppt\/slides\/slide\d+\.xml$/.test(p)).sort();
    const lastSlide = getPartText(pkg, slideParts[slideParts.length - 1]) ?? "";
    expect(lastSlide).toContain("Refilled Title");
  });
});
