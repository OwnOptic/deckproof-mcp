/**
 * Test fixtures built in-code (never checked-in binaries), so they never drift
 * from what the engine emits.
 */

import PptxGenJSImport from "pptxgenjs";

const PptxGenJS = PptxGenJSImport as unknown as {
  new (): {
    layout: string;
    defineSlideMaster(props: Record<string, unknown>): void;
    addSlide(props?: Record<string, unknown>): { addText(t: unknown, o?: Record<string, unknown>): unknown };
    write(props?: { outputType?: string }): Promise<Buffer | Uint8Array>;
  };
};

/**
 * A stand-in "uploaded company template": a branded master with real title +
 * body placeholders and a non-default background color, plus one content slide
 * to clone as a stencil. Returns a Buffer.
 */
export async function buildSyntheticTemplate(): Promise<Buffer> {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.defineSlideMaster({
    title: "BRAND_MASTER",
    background: { color: "0B2A4A" },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.5, w: 12, h: 1.5, color: "FFD166", fontSize: 40 }, text: "Title" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 2.5, w: 12, h: 3, color: "FFFFFF", fontSize: 20 }, text: "Body" } },
    ],
  });
  const slide = pptx.addSlide({ masterName: "BRAND_MASTER" });
  slide.addText("Original Company Title", { placeholder: "title" });
  slide.addText("Original company body content", { placeholder: "body" });
  const out = await pptx.write({ outputType: "nodebuffer" });
  return Buffer.from(out as Uint8Array);
}

/** A declarative slide spec covering all 15 archetypes, for from-scratch tests. */
export const ALL_FIFTEEN_SLIDES = [
  { archetype: "cover", title: "All 15", eyebrow: "Test", subtitle: "sub", context: "ctx" },
  { archetype: "agenda", title: "Agenda", items: [{ label: "Intro", duration: "5m" }, { label: "Body", detail: "meat" }] },
  { archetype: "contentBullets", title: "Points", intro: "why", bullets: ["a", "b", "c"] },
  { archetype: "twoColumns", title: "Compare", left: { heading: "L", bullets: ["x"] }, right: { heading: "R", body: "y" } },
  { archetype: "quote", quote: "Ship it.", author: "Someone", role: "Eng" },
  { archetype: "sectionDivider", title: "Part 2", subtitle: "details" },
  { archetype: "closingNextSteps", title: "Next", steps: [{ title: "Review", owner: "A" }] },
  { archetype: "comparisonTable", title: "Plans", options: ["Free", "Pro"], features: [{ name: "Seats", values: ["1", "10"] }, { name: "SLA", values: [false, true] }] },
  { archetype: "dataTable", title: "Data", headers: ["Q", "Rev"], rows: [["Q1", "100"]] },
  { archetype: "timeline", title: "Roadmap", milestones: [{ when: "Jan", label: "Start" }, { when: "Jun", label: "GA" }] },
  { archetype: "statsBanner", title: "Impact", stats: [{ value: "3x", label: "faster" }, { value: "0", label: "bugs" }] },
  { archetype: "cardGrid", title: "Feats", cards: [{ title: "A", body: "a" }, { title: "B", body: "b" }] },
  { archetype: "orgChart", title: "Team", root: { title: "CEO" }, children: [{ title: "CTO" }, { title: "CFO" }] },
  { archetype: "matrixQuadrant", title: "Matrix", xAxis: "Effort", yAxis: "Value", quadrants: [{ heading: "Q1", items: ["x"] }, { heading: "Q2" }, { heading: "Q3" }, { heading: "Q4" }] },
  { archetype: "verticalSteps", title: "Process", steps: [{ title: "Plan" }, { title: "Ship" }] },
];
