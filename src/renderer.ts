import * as fs from 'fs';
import * as path from 'path';
import { SlideImage } from './types.js';

/**
 * Attempt to render all slides to PNG using PowerPoint COM automation (winax).
 *
 * Requirements (all must be true for rendering to work):
 *   - Windows OS
 *   - Microsoft PowerPoint installed
 *   - winax optional dependency compiled successfully
 *
 * Returns a map of slideNumber → SlideImage.
 * Returns an empty map (silently) if winax is unavailable or rendering fails.
 */
export async function renderSlides(
  pptxAbsPath: string,
  outputDir: string,
): Promise<Map<number, SlideImage>> {
  const rendersDir = path.join(outputDir, 'renders');
  const result = new Map<number, SlideImage>();

  // Dynamic import — winax is optional; don't let a missing module crash the parser
  let winax: { Object: new (progId: string) => unknown };
  try {
    // @ts-ignore — winax is an optional native dependency without types
    winax = (await import('winax')) as typeof winax;
  } catch {
    return result; // not installed or failed to compile — skip rendering
  }

  let pptApp: Record<string, unknown> | null = null;
  let pres: Record<string, unknown> | null = null;

  try {
    fs.mkdirSync(rendersDir, { recursive: true });

    pptApp = new winax.Object('PowerPoint.Application') as Record<string, unknown>;
    (pptApp as unknown as Record<string, boolean>).Visible = false;

    const presentations = pptApp['Presentations'] as Record<string, unknown>;
    pres = (presentations['Open'] as Function)(pptxAbsPath, true /* ReadOnly */) as Record<string, unknown>;

    (pres['Export'] as Function)(rendersDir, 'PNG');

    (pres['Close'] as Function)();
    pres = null;
    (pptApp['Quit'] as Function)();
    pptApp = null;

    // PowerPoint names the exported files 'Slide1.png', 'Slide2.png', etc.
    for (const file of fs.readdirSync(rendersDir)) {
      const match = /^slide(\d+)\.png$/i.exec(file);
      if (match) {
        result.set(parseInt(match[1], 10), {
          mime: 'image/png',
          path: `./renders/${file}`,
        });
      }
    }
  } catch {
    // Rendering failed — clean up COM objects and continue without renders
    if (pres !== null) {
      try { (pres['Close'] as Function)(); } catch { /* ignore */ }
    }
    if (pptApp !== null) {
      try { (pptApp['Quit'] as Function)(); } catch { /* ignore */ }
    }
    // Return empty map — the rest of the output is still valid
  }

  return result;
}
