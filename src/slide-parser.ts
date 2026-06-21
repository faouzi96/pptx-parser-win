import { XMLParser } from 'fast-xml-parser';
import { v4 as uuidv4 } from 'uuid';
import { PptxArchive } from './loader.js';
import { extractTextElements, extractTitle } from './text-extractor.js';
import { extractImageElements, ExtractedImage } from './image-extractor.js';
import { extractTableElements } from './table-extractor.js';
import { Slide, SlideElement, ChartElement, OutputType } from './types.js';
import { parseBounds, extractTextFromTxBody } from './utils.js';

const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** fast-xml-parser instance shared across all slide parsing */
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) =>
    [
      'Relationship',
      'p:sp', 'p:pic', 'p:graphicFrame',
      'a:p', 'a:r',
      'a:tr', 'a:tc',
    ].includes(name),
});

// ---------------------------------------------------------------------------
// Relationship resolution
// ---------------------------------------------------------------------------

/**
 * Load and parse the .rels file for a given slide, returning a map of
 * rId → normalized ZIP path (e.g. 'ppt/media/image1.png').
 */
function resolveSlideRels(archive: PptxArchive, slidePath: string): Map<string, string> {
  // ZIP paths always use forward slashes
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));         // 'ppt/slides'
  const slideFilename = slidePath.substring(slidePath.lastIndexOf('/') + 1);   // 'slide1.xml'
  const relsPath = `${slideDir}/_rels/${slideFilename}.rels`;

  const relsXml = archive.files.get(relsPath);
  if (!relsXml) return new Map();

  const relsDoc = parser.parse(relsXml) as Record<string, unknown>;
  const relationships = ((relsDoc?.['Relationships'] as Record<string, unknown>)?.['Relationship'] as unknown[]) ?? [];

  const result = new Map<string, string>();
  for (const rel of relationships) {
    const r = rel as Record<string, string>;
    const id = r['@_Id'];
    const rawTarget = r['@_Target'] ?? '';
    // Resolve relative paths like '../media/image1.png' against the slide dir
    const parts = `${slideDir}/${rawTarget}`.split('/');
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === '..') resolved.pop();
      else if (part !== '.') resolved.push(part);
    }
    result.set(id, resolved.join('/'));
  }
  return result;
}

// ---------------------------------------------------------------------------
// Notes extraction
// ---------------------------------------------------------------------------

function extractNotes(archive: PptxArchive, rels: Map<string, string>): string | null {
  let notesPath: string | undefined;
  for (const target of rels.values()) {
    if (target.includes('notesSlides/')) {
      notesPath = target;
      break;
    }
  }
  if (!notesPath) return null;

  const notesXml = archive.files.get(notesPath);
  if (!notesXml) return null;

  const notesDoc = parser.parse(notesXml) as Record<string, unknown>;
  const spTree = (
    ((notesDoc?.['p:notes'] as Record<string, unknown>)?.['p:cSld'] as Record<string, unknown>)?.['p:spTree'] as Record<string, unknown>
  );
  if (!spTree) return null;

  const rawShapes = spTree?.['p:sp'];
  const shapes: unknown[] = Array.isArray(rawShapes) ? rawShapes : rawShapes ? [rawShapes] : [];

  const lines: string[] = [];
  for (const sp of shapes) {
    const shape = sp as Record<string, unknown>;
    const nvSpPr = shape?.['p:nvSpPr'] as Record<string, unknown> | undefined;
    const nvPr = nvSpPr?.['p:nvPr'] as Record<string, unknown> | undefined;
    const ph = nvPr?.['p:ph'] as Record<string, string> | undefined;

    // Only the body placeholder (idx=1) contains speaker notes
    // idx=0 is the slide image thumbnail; type='sldImg' is also the slide preview
    const phIdx = ph?.['@_idx'];
    const phType = ph?.['@_type'] ?? '';
    if (phType === 'sldImg') continue;
    if (phIdx !== undefined && String(phIdx) !== '1') continue;

    const txBody = shape?.['p:txBody'] as Record<string, unknown> | undefined;
    if (!txBody) continue;
    const text = extractTextFromTxBody(txBody).trim();
    if (text) lines.push(text);
  }

  return lines.length > 0 ? lines.join('\n') : null;
}

// ---------------------------------------------------------------------------
// Chart extraction (preserve reference only, no semantic reconstruction)
// ---------------------------------------------------------------------------

function extractChartElements(graphicFrameNodes: unknown[]): ChartElement[] {
  const results: ChartElement[] = [];

  for (const frame of graphicFrameNodes) {
    const f = frame as Record<string, unknown>;
    const graphic = f?.['a:graphic'] as Record<string, unknown> | undefined;
    const graphicData = graphic?.['a:graphicData'] as Record<string, unknown> | undefined;
    if (!graphicData) continue;

    const uri = (graphicData['@_uri'] as string) ?? '';
    if (uri !== CHART_URI) continue;

    // c:chart element holds the relationship ID
    const chartNode = graphicData?.['c:chart'] as Record<string, string> | undefined;
    const rId = chartNode?.['@_r:id'] ?? '';

    const xfrm = f?.['p:xfrm'] as Record<string, unknown> | undefined;
    const bounds = xfrm ? parseBounds(xfrm) : { x: 0, y: 0, width: 0, height: 0 };

    results.push({
      id: `el-${uuidv4()}`,
      type: 'chart',
      relationshipId: rId,
      bounds,
      order: 0,
    });
  }

  return results;
}

// ---------------------------------------------------------------------------
// Reading order sort
// ---------------------------------------------------------------------------

function sortByReadingOrder(elements: SlideElement[]): SlideElement[] {
  return elements
    .sort((a, b) => {
      const dy = a.bounds.y - b.bounds.y;
      if (dy !== 0) return dy;
      return a.bounds.x - b.bounds.x;
    })
    .map((el, i) => ({ ...el, order: i + 1 }));
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface ParsedSlide {
  slide: Slide;
  /** Images to be written to disk after all slides are parsed */
  extractedImages: ExtractedImage[];
}

export function parseSlide(
  archive: PptxArchive,
  slidePath: string,
  slideNumber: number,
  outputType: OutputType = 'JSON_AND_MEDIA',
): ParsedSlide {
  const slideXml = archive.files.get(slidePath);
  if (!slideXml) {
    return {
      slide: { slideNumber, title: null, notes: null, fullSlideImage: null, elements: [] },
      extractedImages: [],
    };
  }

  const slideDoc = parser.parse(slideXml) as Record<string, unknown>;
  const sld = slideDoc?.['p:sld'] as Record<string, unknown> | undefined;
  const cSld = sld?.['p:cSld'] as Record<string, unknown> | undefined;
  const spTree = cSld?.['p:spTree'] as Record<string, unknown> | undefined;

  const spNodes: unknown[] = (spTree?.['p:sp'] as unknown[]) ?? [];
  const picNodes: unknown[] = (spTree?.['p:pic'] as unknown[]) ?? [];
  const frameNodes: unknown[] = (spTree?.['p:graphicFrame'] as unknown[]) ?? [];

  const rels = resolveSlideRels(archive, slidePath);

  // Extract all element types
  const textElements = extractTextElements(spNodes);
  const extractedImages = extractImageElements(picNodes, rels, archive, outputType);
  const tableElements = extractTableElements(frameNodes);
  const chartElements = extractChartElements(frameNodes);

  const allElements: SlideElement[] = sortByReadingOrder([
    ...textElements,
    ...extractedImages.map((e) => e.element),
    ...tableElements,
    ...chartElements,
  ]);

  return {
    slide: {
      slideNumber,
      title: extractTitle(spNodes),
      notes: extractNotes(archive, rels),
      fullSlideImage: null, // populated later by renderer if available
      elements: allElements,
    },
    extractedImages,
  };
}
