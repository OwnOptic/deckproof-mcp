/**
 * Minimal local type shim for the small slice of the pptxgenjs API this
 * engine actually uses.
 *
 * pptxgenjs ships an older UMD-style `.d.ts` (`export as namespace X` +
 * unexported `declare class`/`declare namespace` merge) that TypeScript's
 * Node16 module resolution does not resolve cleanly - the default import's
 * static type collapses to the module's ambient namespace object instead of
 * the declared class, so `new PptxGenJS()` and `PptxGenJS.Slide` both fail
 * to type-check even though the runtime value is correct (esModuleInterop
 * handles that part fine). Rather than fight upstream declaration-merging
 * quirks, we declare our own narrow, accurate types for the handful of
 * methods used here and cast the import to them once, at the boundary.
 */

export interface PptxSlide {
  addText(text: unknown, options?: Record<string, unknown>): PptxSlide;
  addShape(shapeType: unknown, options?: Record<string, unknown>): PptxSlide;
  addTable(rows: unknown[], options?: Record<string, unknown>): PptxSlide;
  addImage(options: Record<string, unknown>): PptxSlide;
}

export interface PptxPresentation {
  layout: string;
  title: string;
  defineSlideMaster(props: Record<string, unknown>): void;
  addSlide(props?: Record<string, unknown>): PptxSlide;
  write(props?: { outputType?: string }): Promise<Buffer | Uint8Array | ArrayBuffer | string>;
}

export type PptxGenJSConstructor = new () => PptxPresentation;
