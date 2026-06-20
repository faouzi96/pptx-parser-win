import { fileURLToPath } from 'url';
import { parsePptx } from './src/pptxParse.js';
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
} from './src/types.js';
// ---------------------------------------------------------------------------
// CLI entry point
// Usage: node dist/index.js <pptx-path> <output-dir>
// ---------------------------------------------------------------------------
// CLI entry point
const isMain = process.argv[1] === fileURLToPath(import.meta.url);
if (isMain) {
  const [, , pptxArg, outputArg, imgOutputArg] = process.argv;
  let isValidWithImgOutput = true;

  if (imgOutputArg) {
    if (imgOutputArg === 'true' || imgOutputArg === 'false') {
      isValidWithImgOutput = imgOutputArg === 'true';
    }
  }

  if (!pptxArg) {
    process.stderr.write(
      "Usage: node dist/index.js <pptx-path> [output-dir] [with_img_output=true|false]\n",
    );
    process.exit(1);
  }

  parsePptx(pptxArg, outputArg, isValidWithImgOutput)
    .then((presentation) => {
      process.stdout.write(
        JSON.stringify({
          success: true,
          slideCount: presentation.slides.length,
          outputDir: outputArg,
        }) + "\n",
      );
    })
    .catch((err: Error) => {
      process.stderr.write(
        JSON.stringify({ success: false, error: err.message }) + "\n",
      );
      process.exit(1);
    });
}