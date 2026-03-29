import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

import { openPackage } from '@ooxml/core';
import { parsePptx } from '@ooxml/pptx';

const corpusRoot = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(os.homedir(), 'Desktop', 'ppt-samples');
const outputPath = process.argv[3]
  ? path.resolve(process.argv[3])
  : path.join(process.cwd(), 'benchmarks', 'reports', 'latest-ppt-sample-report.json');

async function listSampleDirectories(root) {
  const entries = await readdir(root, { withFileTypes: true });
  return entries.filter((entry) => entry.isDirectory()).map((entry) => path.join(root, entry.name)).sort();
}

async function summarizeSampleDirectory(sampleDirectory) {
  const entries = await readdir(sampleDirectory, { withFileTypes: true });
  const pptxFiles = entries.filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith('.pptx'));
  const pngFiles = entries.filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith('.png'));
  const sampleName = path.basename(sampleDirectory);

  if (!pptxFiles.length) {
    return {
      sample: sampleName,
      sampleDirectory,
      pptxCount: 0,
      pngCount: pngFiles.length,
      error: 'No PPTX file found in sample directory.'
    };
  }

  const pptxPath = path.join(sampleDirectory, pptxFiles[0].name);
  const bytes = new Uint8Array(await readFile(pptxPath));
  const packageGraph = await openPackage(bytes);
  const presentation = parsePptx(packageGraph);

  const imageShapeCount = presentation.slides.reduce((sum, slide) => sum + slide.shapes.filter((shape) => shape.media?.type === 'image').length, 0);
  const embeddedObjectCount = presentation.slides.reduce((sum, slide) => sum + slide.shapes.filter((shape) => shape.media?.type === 'embeddedObject').length, 0);
  const textShapeCount = presentation.slides.reduce((sum, slide) => sum + slide.shapes.filter((shape) => !shape.media).length, 0);

  return {
    sample: sampleName,
    sampleDirectory,
    pptxPath,
    pptxCount: pptxFiles.length,
    pngCount: pngFiles.length,
    parsedSlideCount: presentation.slides.length,
    officeDocumentKind: packageGraph.officeDocumentKind,
    rootDocumentUri: packageGraph.rootDocumentUri,
    deckSize: presentation.size,
    imageShapeCount,
    embeddedObjectCount,
    textShapeCount,
    firstSlideTitle: presentation.slides[0]?.title ?? null,
    diagnostics: packageGraph.diagnostics
  };
}

const samples = await listSampleDirectories(corpusRoot);
const results = [];

for (const sampleDirectory of samples) {
  try {
    results.push(await summarizeSampleDirectory(sampleDirectory));
  } catch (error) {
    results.push({
      sample: path.basename(sampleDirectory),
      sampleDirectory,
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

await mkdir(path.dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify({
  generatedAt: new Date().toISOString(),
  corpusRoot,
  sampleCount: samples.length,
  results
}, null, 2));

console.log(JSON.stringify({
  outputPath: path.relative(process.cwd(), outputPath).replaceAll('\\', '/'),
  corpusRoot,
  sampleCount: samples.length
}, null, 2));
