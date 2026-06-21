/**
 * Parses CLI arguments. Supports both modern flags and legacy positional formats.
 */
export function parseArgs(args: string[]): { pptxPath?: string; outputPath?: string; outputType?: string } {
  const hasFlags = args.some(arg => arg.startsWith('--') || arg.startsWith('-'));

  if (hasFlags) {
    const result: { pptxPath?: string; outputPath?: string; outputType?: string } = {};
    for (let i = 0; i < args.length; i++) {
      const arg = args[i];
      if (arg === '--pptx-path' || arg === '-p') {
        result.pptxPath = args[i + 1];
        i++;
      } else if (arg === '--output-path' || arg === '-o') {
        result.outputPath = args[i + 1];
        i++;
      } else if (arg === '--output-type' || arg === '-t') {
        result.outputType = args[i + 1];
        i++;
      }
    }
    return result;
  }

  return {
    pptxPath: args[0],
    outputPath: args[1],
    outputType: args[2],
  };
}
