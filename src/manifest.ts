import { XMLParser } from 'fast-xml-parser';
import { PptxArchive } from './loader.js';

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => ['Relationship', 'p:sldId'].includes(name),
});

export interface SlideManifest {
  /** Ordered list of ZIP paths to slide XML files, e.g. 'ppt/slides/slide1.xml' */
  slidePaths: string[];
  presentationTitle: string | null;
}

export function parseManifest(archive: PptxArchive): SlideManifest {
  const presXml = archive.files.get('ppt/presentation.xml');
  if (!presXml) throw new Error('Invalid PPTX: missing ppt/presentation.xml');

  const presRelsXml = archive.files.get('ppt/_rels/presentation.xml.rels');
  if (!presRelsXml) throw new Error('Invalid PPTX: missing ppt/_rels/presentation.xml.rels');

  const presDoc = parser.parse(presXml) as Record<string, unknown>;
  const relsDoc = parser.parse(presRelsXml) as Record<string, unknown>;

  // Build rId → slide target map from the rels file
  const relationships = ((relsDoc?.['Relationships'] as Record<string, unknown>)?.['Relationship'] as unknown[]) ?? [];
  const rIdToTarget = new Map<string, string>();
  for (const rel of relationships) {
    const r = rel as Record<string, string>;
    if ((r['@_Type'] ?? '').endsWith('/slide')) {
      rIdToTarget.set(r['@_Id'], r['@_Target']);
    }
  }

  // Get the ordered slide ID list from presentation.xml
  const pres = presDoc?.['p:presentation'] as Record<string, unknown> | undefined;
  const sldIdLst = pres?.['p:sldIdLst'] as Record<string, unknown> | undefined;
  const sldIdList = (sldIdLst?.['p:sldId'] as unknown[]) ?? [];

  const slidePaths: string[] = [];
  for (const sldId of sldIdList) {
    const s = sldId as Record<string, string>;
    const rId = s['@_r:id'] ?? '';
    const target = rIdToTarget.get(rId);
    if (!target) continue;

    // Target is relative to ppt/, e.g. 'slides/slide1.xml'
    // Some files may use an absolute ZIP path like '/ppt/slides/slide1.xml'
    if (target.startsWith('/')) {
      slidePaths.push(target.slice(1));
    } else {
      slidePaths.push(`ppt/${target.replace(/^\.\//, '')}`);
    }
  }

  return { slidePaths, presentationTitle: null };
}
