import { Bounds } from './types.js';

/** Convert EMU (English Metric Units) to pixels at 96 DPI. Standard slide: 9144000×6858000 EMU = 960×720px */
export function emuToPx(emu: number): number {
  return Math.round(emu / 9525);
}

/** Parse an <a:xfrm> XML node into a Bounds object */
export function parseBounds(xfrmNode: Record<string, unknown>): Bounds {
  const off = (xfrmNode?.['a:off'] ?? {}) as Record<string, string>;
  const ext = (xfrmNode?.['a:ext'] ?? {}) as Record<string, string>;
  return {
    x: emuToPx(parseInt(off['@_x'] ?? '0', 10)),
    y: emuToPx(parseInt(off['@_y'] ?? '0', 10)),
    width: emuToPx(parseInt(ext['@_cx'] ?? '0', 10)),
    height: emuToPx(parseInt(ext['@_cy'] ?? '0', 10)),
  };
}

/** Extract all visible text from a <p:txBody> node */
export function extractTextFromTxBody(txBody: Record<string, unknown>): string {
  const rawParas = txBody?.['a:p'];
  const paragraphs: unknown[] = Array.isArray(rawParas) ? rawParas : rawParas ? [rawParas] : [];

  const lines: string[] = [];
  for (const para of paragraphs) {
    const p = para as Record<string, unknown>;
    const rawRuns = p?.['a:r'];
    const runs: unknown[] = Array.isArray(rawRuns) ? rawRuns : rawRuns ? [rawRuns] : [];
    const line = runs.map((r) => ((r as Record<string, unknown>)?.['a:t'] as string) ?? '').join('');
    if (line.trim()) lines.push(line);
  }

  return lines.join('\n');
}
