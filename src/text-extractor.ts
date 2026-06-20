import { v4 as uuidv4 } from 'uuid';
import { TextElement } from './types.js';
import { parseBounds, extractTextFromTxBody } from './utils.js';

/**
 * Extract TextElement[] from an array of <p:sp> shape nodes.
 * Each shape with a non-empty txBody becomes a TextElement.
 */
export function extractTextElements(spNodes: unknown[]): TextElement[] {
  const results: TextElement[] = [];

  for (const sp of spNodes) {
    const shape = sp as Record<string, unknown>;
    const txBody = shape?.['p:txBody'] as Record<string, unknown> | undefined;
    if (!txBody) continue;

    const content = extractTextFromTxBody(txBody);
    if (!content.trim()) continue;

    const spPr = shape?.['p:spPr'] as Record<string, unknown> | undefined;
    const xfrm = spPr?.['a:xfrm'] as Record<string, unknown> | undefined;
    const bounds = xfrm ? parseBounds(xfrm) : { x: 0, y: 0, width: 0, height: 0 };

    results.push({
      id: `el-${uuidv4()}`,
      type: 'text',
      content: content.trim(),
      bounds,
      order: 0, // sorted later by reading order
    });
  }

  return results;
}

/**
 * Determine the slide title from its shape list.
 * Title placeholders have ph type="title" or type="ctrTitle".
 */
export function extractTitle(spNodes: unknown[]): string | null {
  for (const sp of spNodes) {
    const shape = sp as Record<string, unknown>;
    const nvSpPr = shape?.['p:nvSpPr'] as Record<string, unknown> | undefined;
    const nvPr = nvSpPr?.['p:nvPr'] as Record<string, unknown> | undefined;
    const ph = nvPr?.['p:ph'] as Record<string, string> | undefined;
    const phType = ph?.['@_type'] ?? '';

    if (phType === 'title' || phType === 'ctrTitle') {
      const txBody = shape?.['p:txBody'] as Record<string, unknown> | undefined;
      if (!txBody) continue;
      const text = extractTextFromTxBody(txBody).trim();
      if (text) return text;
    }
  }
  return null;
}
