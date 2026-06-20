import { v4 as uuidv4 } from 'uuid';
import { TableElement } from './types.js';
import { parseBounds } from './utils.js';

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

function extractCellText(tc: Record<string, unknown>): string {
  const txBody = tc?.['a:txBody'] as Record<string, unknown> | undefined;
  if (!txBody) return '';

  const rawParas = txBody?.['a:p'];
  const paragraphs: unknown[] = Array.isArray(rawParas) ? rawParas : rawParas ? [rawParas] : [];

  return paragraphs
    .map((p) => {
      const para = p as Record<string, unknown>;
      const rawRuns = para?.['a:r'];
      const runs: unknown[] = Array.isArray(rawRuns) ? rawRuns : rawRuns ? [rawRuns] : [];
      return runs.map((r) => ((r as Record<string, unknown>)?.['a:t'] as string) ?? '').join('');
    })
    .join(' ')
    .trim();
}

/**
 * Extract TableElement[] from <p:graphicFrame> nodes that contain an <a:tbl>.
 */
export function extractTableElements(graphicFrameNodes: unknown[]): TableElement[] {
  const results: TableElement[] = [];

  for (const frame of graphicFrameNodes) {
    const f = frame as Record<string, unknown>;
    const graphic = f?.['a:graphic'] as Record<string, unknown> | undefined;
    const graphicData = graphic?.['a:graphicData'] as Record<string, unknown> | undefined;
    if (!graphicData) continue;

    const uri = (graphicData['@_uri'] as string) ?? '';
    if (uri !== TABLE_URI) continue;

    const tbl = graphicData?.['a:tbl'] as Record<string, unknown> | undefined;
    if (!tbl) continue;

    const rawRows = tbl?.['a:tr'];
    const tableRows: unknown[] = Array.isArray(rawRows) ? rawRows : rawRows ? [rawRows] : [];

    const rows: string[][] = tableRows.map((tr) => {
      const row = tr as Record<string, unknown>;
      const rawCells = row?.['a:tc'];
      const cells: unknown[] = Array.isArray(rawCells) ? rawCells : rawCells ? [rawCells] : [];
      return cells.map((tc) => extractCellText(tc as Record<string, unknown>));
    });

    // Tables in graphicFrame use <p:xfrm> not <p:spPr>/<a:xfrm>
    const xfrm = f?.['p:xfrm'] as Record<string, unknown> | undefined;
    const bounds = xfrm ? parseBounds(xfrm) : { x: 0, y: 0, width: 0, height: 0 };

    results.push({
      id: `el-${uuidv4()}`,
      type: 'table',
      rows,
      bounds,
      order: 0,
    });
  }

  return results;
}
