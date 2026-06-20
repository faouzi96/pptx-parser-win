import * as fs from 'fs';
import * as path from 'path';
import { Presentation } from './types.js';

/**
 * Serialize and write the final Presentation object to outputDir/presentation.json.
 */
export function writePresentation(presentation: Presentation, outputDir: string): void {
  fs.mkdirSync(outputDir, { recursive: true });
  const outPath = path.join(outputDir, 'presentation.json');
  fs.writeFileSync(outPath, JSON.stringify(presentation, null, 2), 'utf-8');
}
