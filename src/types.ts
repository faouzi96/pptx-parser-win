/**
 * All shared TypeScript types for the PPTX parser output.
 * These map 1:1 to the JSON written to presentation.json.
 */

/** Pixel coordinates/dimensions (converted from EMU at 96 DPI) */
export interface Bounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TextElement {
  id: string;
  type: 'text';
  content: string;
  bounds: Bounds;
  order: number;
}

export interface ImageElement {
  id: string;
  type: 'image';
  /** MIME type, e.g. 'image/png' */
  mime: string;
  /** Path relative to outputDir, e.g. './media/image1.png' */
  path?: string;
  base64?: string;
  bounds: Bounds;
  order: number;
}

export interface TableElement {
  id: string;
  type: 'table';
  /** Row-major 2D array of cell text values */
  rows: string[][];
  bounds: Bounds;
  order: number;
}

export interface ChartElement {
  id: string;
  type: 'chart';
  /** The slide relationship ID referencing the chart data */
  relationshipId: string;
  bounds: Bounds;
  order: number;
}

export type SlideElement = TextElement | ImageElement | TableElement | ChartElement;

export interface SlideImage {
  mime: string;
  /** Path relative to outputDir, e.g. './renders/slide1.png' */
  path: string;
}

export interface Slide {
  slideNumber: number;
  title: string | null;
  notes: string | null;
  /** Full-slide PNG render — null if PowerPoint rendering was unavailable */
  fullSlideImage: SlideImage | null;
  elements: SlideElement[];
}

export interface Presentation {
  title: string | null;
  slides: Slide[];
}

export type OutputType = 'JSON_AND_MEDIA' | 'JSON_EMBEDDED' | 'NONE';