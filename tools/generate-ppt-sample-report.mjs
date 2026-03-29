import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';

const root = process.cwd();
const sampleRoot = process.env.PPT_SAMPLE_ROOT ?? path.join(process.env.HOME ?? '', 'Desktop', 'ppt-samples');
const outputPath = path.join(root, 'benchmarks', 'reports', 'ppt-sample-report.json');

const entries = await readdir(sampleRoot, { withFileTypes: true }).catch(() => []);
const results = [];

for (const entry of entries) {
  if (!entry.isDirectory()) {
    continue;
  }

  const folder = path.join(sampleRoot, entry.name);
  const pptxPath = path.join(folder, 'sample.pptx');
  const pngEntries = (await readdir(folder, { withFileTypes: true }).catch(() => []))
    .filter((dirent) => dirent.isFile() && dirent.name.endsWith('.png'))
    .map((dirent) => dirent.name)
    .sort();

  try {
    const bytes = new Uint8Array(await readFile(pptxPath));
    const packageGraph = await openPackage(bytes);
    const presentation = packageGraph.officeDocumentKind === 'pptx' ? parsePptx(packageGraph) : null;

    results.push({
      sample: entry.name,
      pptxPath: path.relative(root, pptxPath).replaceAll('\\', '/'),
      pngCount: pngEntries.length,
      rootDocumentUri: packageGraph.rootDocumentUri,
      officeDocumentKind: packageGraph.officeDocumentKind,
      slideCount: presentation?.slides.length ?? 0,
      slideSize: presentation?.size ?? null,
      diagnostics: packageGraph.diagnostics,
      parseStatus: presentation ? 'parsed' : 'unsupported'
    });
  } catch (error) {
    results.push({
      sample: entry.name,
      pptxPath: path.relative(root, pptxPath).replaceAll('\\', '/'),
      pngCount: pngEntries.length,
      rootDocumentUri: null,
      officeDocumentKind: 'unknown',
      slideCount: 0,
      slideSize: null,
      diagnostics: [],
      parseStatus: 'failed',
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

await mkdir(path.dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify({
  generatedAt: new Date().toISOString(),
  sampleRoot,
  sampleCount: results.length,
  results
}, null, 2));

console.log(JSON.stringify({
  outputPath: path.relative(root, outputPath).replaceAll('\\', '/'),
  sampleCount: results.length
}, null, 2));
