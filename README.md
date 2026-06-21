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

// 1. Write everything (JSON + media files on disk)
const presentationFiles = await parsePptx('presentation.pptx', './output', 'JSON_AND_MEDIA');

// 2. Write JSON metadata only (no image files on disk, images are base64-embedded)
const presentationJsonOnly = await parsePptx('presentation.pptx', './output', 'JSON_EMBEDDED');

// 3. Write absolutely nothing to disk (keep everything in memory, images are base64-embedded)
const presentationInMemory = await parsePptx('presentation.pptx', undefined, 'NONE');
```

Run as a CLI:

```bash
npx pptx-parser presentation.pptx ./output JSON_AND_MEDIA
npx pptx-parser presentation.pptx ./output JSON_EMBEDDED
npx pptx-parser presentation.pptx ./output NONE
```

### CLI arguments

- `<pptx-path>` — required input file path
- `[output-dir]` — optional target folder for output files
- `[output-mode]` — optional output mode (`JSON_AND_MEDIA` | `JSON_EMBEDDED` | `NONE`), defaults to `JSON_AND_MEDIA`.

### Output behavior

- **`JSON_AND_MEDIA`** (default): Extracted media files are written to `output/media/`, slide renders to `output/renders/`, and `presentation.json` is written to the output directory. Image elements in JSON contain a `path` property.
- **`JSON_EMBEDDED`**: Only `presentation.json` is written to the output directory (no image files are created). Image elements in JSON contain an embedded `base64` property.
- **`NONE`**: No files are written to disk at all. Image elements in the returned `Presentation` object contain an embedded `base64` property.

## Package Contents

- `main`: `dist/index.js`
- `bin`: `pptx-parser`
- exports root entrypoint and CLI from the same file
- compiled files are published from `dist/`
