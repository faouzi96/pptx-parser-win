import * as fs from "fs";
import * as path from "path";

import { loadPptx } from "./loader.js";
import { parseManifest } from "./manifest.js";
import { parseSlide } from "./slide-parser.js";
import { writeMediaFiles, ExtractedImage } from "./image-extractor.js";
import { renderSlides } from "./renderer.js";
import { writePresentation } from "./normalizer.js";
import { Presentation, Slide, OutputType } from "./types.js";

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
  OutputType,
} from "./types.js";

/**
 * Parse a .pptx file into a structured Presentation object.
 *
 * Side effects (written to outputDir or current directory):
 *   - 'JSON_AND_MEDIA' mode: writes outputDir/media/, outputDir/renders/ and outputDir/presentation.json
 *   - 'JSON_EMBEDDED' mode: writes outputDir/presentation.json only
 *   - 'NONE' mode: writes no files to disk
 *
 * @param pptxPathArg  Absolute or relative path to the .pptx file
 * @param outputDirArg Optional directory to write output files into (defaults to process.cwd())
 * @param outputType   Output behavior mode: 'JSON_AND_MEDIA' (default), 'JSON_EMBEDDED', or 'NONE'.
 * @returns            The parsed Presentation object
 */
export async function parsePptx(
  pptxPathArg: string,
  outputDirArg?: string,
  outputType: OutputType = 'JSON_AND_MEDIA',
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
      outputType,
    );
    slides.push(slide);
    allExtractedImages.push(...extractedImages);
  }

  // 4. Write extracted media files to outputDir/media/
  if (outputType === 'JSON_AND_MEDIA') {
    writeMediaFiles(allExtractedImages, archive, path.join(outputPath, "media"));
  }

  // 5. Attempt full-slide rendering via PowerPoint COM (optional, Windows + PPT only)
  if (outputType === 'JSON_AND_MEDIA') {
    const renders = await renderSlides(absPath, outputPath);
    for (const slide of slides) {
      const render = renders.get(slide.slideNumber);
      if (render) slide.fullSlideImage = render;
    }
  }

  // 6. Assemble and write the final JSON
  const presentation: Presentation = {
    title: manifest.presentationTitle,
    slides,
  };
  if (outputType !== 'NONE') {
    writePresentation(presentation, outputPath);
  }

  return presentation;
}
