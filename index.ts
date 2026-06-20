import * as fs from 'fs';
import * as path from 'path';

import { loadPptx } from './src/loader.js';
import { parseManifest } from './src/manifest.js';
import { parseSlide } from './src/slide-parser.js';
import { writeMediaFiles, ExtractedImage } from './src/image-extractor.js';
import { renderSlides } from './src/renderer.js';
import { writePresentation } from './src/normalizer.js';
import { Presentation, Slide } from './src/types.js';
import { fileURLToPath } from 'url';

// Re-export all public types so consumers can import them from this module
export type {
  Presentation,
  Slide,
  SlideElement,
  TextElement,
  ImageElement,
  TableElement,
  ChartElement,
  Bounds,
  SlideImage,
} from './src/types.js';

/**
 * Parse a .pptx file into a structured Presentation object.
 *
 * Side effects (all written to outputDir):
 *   outputDir/media/         — extracted embedded images
 *   outputDir/renders/       — full-slide PNG renders (only if PowerPoint is available)
 *   outputDir/presentation.json — the final normalized output
 *
 * @param pptxPath  Absolute or relative path to the .pptx file
 * @param outputDir Directory to write output files into (created if missing)
 * @returns         The parsed Presentation object
 */
export async function parsePptx(pptxPathArg: string, outputDirArg?: string): Promise<Presentation> {

  const pptxPath=path.normalize(pptxPathArg); 
  const outputPath=outputDirArg ? path.normalize(outputDirArg) : process.cwd();


  const absPath = path.resolve(pptxPath);
  if (!fs.existsSync(absPath)) {
    throw new Error(`File not found: ${absPath}`);
  }

  // 1. Load and unzip the archive into memory
  const archive = await loadPptx(absPath);

  // 2. Resolve the ordered list of slide paths from the presentation manifest
  const manifest = parseManifest(archive);

  // 3. Parse each slide (text, images, tables, charts, notes, title)
  const slides: Slide[] = [];
  const allExtractedImages: ExtractedImage[] = [];

  for (let i = 0; i < manifest.slidePaths.length; i++) {
    const { slide, extractedImages } = parseSlide(
      archive,
      manifest.slidePaths[i],
      i + 1,
    );
    slides.push(slide);
    allExtractedImages.push(...extractedImages);
  }

  // 4. Write extracted media files to outputDir/media/
  writeMediaFiles(allExtractedImages, archive, path.join(outputPath, 'media'));

  // 5. Attempt full-slide rendering via PowerPoint COM (optional, Windows + PPT only)
  const renders = await renderSlides(absPath, outputPath);
  for (const slide of slides) {
    const render = renders.get(slide.slideNumber);
    if (render) slide.fullSlideImage = render;
  }

  // 6. Assemble and write the final JSON
  const presentation: Presentation = {
    title: manifest.presentationTitle,
    slides,
  };
  writePresentation(presentation, outputPath);

  return presentation;
}

// ---------------------------------------------------------------------------
// CLI entry point
// Usage: node dist/index.js <pptx-path> <output-dir>
// ---------------------------------------------------------------------------
// CLI entry point
const isMain = process.argv[1] === fileURLToPath(import.meta.url);
if (isMain) {
  const [, , pptxArg, outputArg] = process.argv;

  if (!pptxArg) {
    process.stderr.write('Usage: node dist/index.js <pptx-path> [OPTIONAL]<output-dir>\n');
    process.exit(1);
  }

  parsePptx(pptxArg, outputArg)
    .then((presentation) => {
      process.stdout.write(
        JSON.stringify({ success: true, slideCount: presentation.slides.length, outputDir: outputArg }) + '\n',
      );
    })
    .catch((err: Error) => {
      process.stderr.write(JSON.stringify({ success: false, error: err.message }) + '\n');
      process.exit(1);
    });
}
