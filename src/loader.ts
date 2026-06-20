import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';

export interface PptxArchive {
  /** All XML and .rels files, keyed by their ZIP path (forward-slash separated) */
  files: Map<string, string>;
  /** All binary media files, keyed by their ZIP path */
  media: Map<string, Buffer>;
}

const XML_EXTENSIONS = new Set(['.xml', '.rels']);
const MEDIA_EXTENSIONS = new Set([
  '.png', '.jpg', '.jpeg', '.gif', '.bmp',
  '.tiff', '.tif', '.webp', '.svg', '.wmf', '.emf',
]);

export async function loadPptx(filePath: string): Promise<PptxArchive> {
  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);

  const files = new Map<string, string>();
  const media = new Map<string, Buffer>();

  const tasks = Object.entries(zip.files).map(async ([name, file]) => {
    if (file.dir) return;
    const ext = path.extname(name).toLowerCase();
    if (XML_EXTENSIONS.has(ext)) {
      files.set(name, await file.async('string'));
    } else if (MEDIA_EXTENSIONS.has(ext)) {
      const ab = await file.async('arraybuffer');
      media.set(name, Buffer.from(ab));
    }
  });

  await Promise.all(tasks);
  return { files, media };
}
