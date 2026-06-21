import * as fs from "fs";
import * as path from "path";
import { v4 as uuidv4 } from "uuid";
import { ImageElement, OutputType } from "./types.js";
import { parseBounds } from "./utils.js";
import { PptxArchive } from "./loader.js";

const MIME_MAP: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".bmp": "image/bmp",
  ".webp": "image/webp",
  ".svg": "image/svg+xml",
  ".wmf": "image/x-wmf",
  ".emf": "image/x-emf",
  ".tiff": "image/tiff",
  ".tif": "image/tiff",
};

export interface ExtractedImage {
  element: ImageElement;
  /** ZIP-internal path to the source media file, e.g. 'ppt/media/image1.png' */
  srcZipPath: string;
}

/**
 * Extract ImageElement[] from <p:pic> nodes.
 * @param picNodes   Array of <p:pic> XML nodes from a slide's spTree
 * @param rels       Map of rId → resolved ZIP path for this slide
 * @param archive    The PPTX ZIP archive
 * @param outputType Output behavior mode
 */
export function extractImageElements(
  picNodes: unknown[],
  rels: Map<string, string>,
  archive: PptxArchive,
  outputType: OutputType = 'JSON_AND_MEDIA',
): ExtractedImage[] {
  const results: ExtractedImage[] = [];

  for (const pic of picNodes) {
    const p = pic as Record<string, unknown>;
    const blipFill = p?.["p:blipFill"] as Record<string, unknown> | undefined;
    const blip = blipFill?.["a:blip"] as Record<string, string> | undefined;
    const rId = blip?.["@_r:embed"] ?? "";

    const srcZipPath = rels.get(rId);
    if (!srcZipPath) continue;

    const buf = archive.media.get(srcZipPath);
    if (!buf) continue;

    const ext = path.extname(srcZipPath).toLowerCase();
    const mime = MIME_MAP[ext] ?? "application/octet-stream";
    const filename = path.basename(srcZipPath);

    const spPr = p?.["p:spPr"] as Record<string, unknown> | undefined;
    const xfrm = spPr?.["a:xfrm"] as Record<string, unknown> | undefined;
    const bounds = xfrm
      ? parseBounds(xfrm)
      : { x: 0, y: 0, width: 0, height: 0 };

    const element: ImageElement = {
      id: `el-${uuidv4()}`,
      type: "image",
      mime,
      bounds,
      order: 0,
    };

    if (outputType === 'JSON_AND_MEDIA') {
      element.path = `./media/${filename}`;
    } else {
      element.base64 = buf.toString("base64");
    }

    results.push({
      element,
      srcZipPath,
    });
  }

  return results;
}

/**
 * Copy extracted media buffers from the archive to outputMediaDir on disk.
 */
export function writeMediaFiles(
  extracted: ExtractedImage[],
  archive: PptxArchive,
  outputMediaDir: string,
): void {
  if (extracted.length === 0) return;
  fs.mkdirSync(outputMediaDir, { recursive: true });

  for (const { srcZipPath, element } of extracted) {
    const buf = archive.media.get(srcZipPath);
    if (buf && element.path) {
      const filename = path.basename(element.path);
      fs.writeFileSync(path.join(outputMediaDir, filename), buf);
    }
  }
}
