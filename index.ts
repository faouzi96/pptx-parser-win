#!/usr/bin/env node
import { fileURLToPath } from 'url';
import { parsePptx } from './src/pptxParse.js';
import { OutputType } from './src/types.js';
import { parseArgs } from './src/cli.js';
export { parsePptx };
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
} from './src/types.js';
// ---------------------------------------------------------------------------
// CLI entry point
// ---------------------------------------------------------------------------

const isMain = process.argv[1] === fileURLToPath(import.meta.url);
if (isMain) {
  const parsed = parseArgs(process.argv.slice(2));
  const pptxArg = parsed.pptxPath;
  const outputArg = parsed.outputPath;
  const imgOutputArg = parsed.outputType;

  let outputType: OutputType = 'JSON_AND_MEDIA';

  if (imgOutputArg) {
    const upper = imgOutputArg.toUpperCase();
    if (upper === 'JSON_AND_MEDIA' || upper === 'JSON_EMBEDDED' || upper === 'NONE') {
      outputType = upper as OutputType;
    } else {
      process.stderr.write(`Error: Invalid output type "${imgOutputArg}". Expected JSON_AND_MEDIA, JSON_EMBEDDED, or NONE.\n`);
      process.exit(1);
    }
  }

  if (!pptxArg) {
    process.stderr.write(
      "Usage:\n" +
      "  pptx-parser --pptx-path <path> [--output-path <path>] [--output-type JSON_AND_MEDIA|JSON_EMBEDDED|NONE]\n" +
      "  OR (positional):\n" +
      "  pptx-parser <pptx-path> [output-dir] [output-mode=JSON_AND_MEDIA|JSON_EMBEDDED|NONE]\n",
    );
    process.exit(1);
  }

  parsePptx(pptxArg, outputArg, outputType)
    .then((presentation) => {
      if (outputType === 'NONE') {
        process.stdout.write(JSON.stringify(presentation, null, 2) + "\n");
      } else {
        process.stdout.write(
          JSON.stringify({
            success: true,
            slideCount: presentation.slides.length,
            outputDir: outputArg,
          }) + "\n",
        );
      }
    })
    .catch((err: Error) => {
      process.stderr.write(
        JSON.stringify({ success: false, error: err.message }) + "\n",
      );
      process.exit(1);
    });
}