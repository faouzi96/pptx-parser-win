import * as fs from "fs";
import * as path from "path";

import { loadPptx } from "./loader.js";
import { parseManifest } from "./manifest.js";
import { parseSlide } from "./slide-parser.js";
import { writeMediaFiles, ExtractedImage } from "./image-extractor.js";
import { renderSlides } from "./renderer.js";
import { writePresentation } from "./normalizer.js";
import { Presentation, Slide } from "./types.js";

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
} from "./types.js";

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
 * @param with_img_output  Whether to write extracted image output and presentation JSON. Defaults to true.
 * @returns         The parsed Presentation object
 */
export async function parsePptx(
  pptxPathArg: string,
  outputDirArg?: string,
  with_img_output = true,
): Promise<Presentation> {
  const pptxPath = path.normalize(pptxPathArg);
  const outputPath = outputDirArg
    ? path.normalize(outputDirArg)
    : process.cwd();

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
  writeMediaFiles(allExtractedImages, archive, path.join(outputPath, "media"));

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
  if (with_img_output) {
    writePresentation(presentation, outputPath);
  }

  return presentation;
}
