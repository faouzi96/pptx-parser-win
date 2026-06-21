# @fz-labs/pptx-parser-win

Deterministic multimodal PPTX parser that converts PowerPoint files to structured JSON without AI.

Repository: https://github.com/faouzi96/pptx-parser-win

## Usage

Install:

```bash
npm install @fz-labs/pptx-parser-win
```

Use as a library:

```ts
import { parsePptx } from '@fz-labs/pptx-parser-win';

const presentation = await parsePptx('presentation.pptx', './output');
// or explicitly control image output:
const presentationWithFiles = await parsePptx('presentation.pptx', './output', true);
const presentationWithoutFiles = await parsePptx('presentation.pptx', './output', false);
```

Run as a CLI:

```bash
npx pptx-parser presentation.pptx ./output
npx pptx-parser presentation.pptx ./output false
```

### CLI arguments

- `<pptx-path>` — required input file path
- `[output-dir]` — optional target folder for output files
- `[with_img_output=true|false]` — optional boolean flag, defaults to `true`

### Output behavior

- When `with_img_output` is `true` (default), extracted media files are written to `output/media/` and `presentation.json` is written to the output directory.
- When `with_img_output` is `false`, no files are written to disk, but the parsed `Presentation` object is still returned.

## Package Contents

- `main`: `dist/index.js`
- `bin`: `pptx-parser`
- exports root entrypoint and CLI from the same file
- compiled files are published from `dist/`
