import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx } from '@ooxml/xlsx';

const root = process.cwd();
const manifestsRoot = path.join(root, 'fixtures', 'manifests');
const resultsOutputPath = path.join(root, 'benchmarks', 'reports', 'latest-fixture-results.json');

async function walk(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) files.push(...await walk(fullPath));
    else if (entry.isFile() && entry.name.endsWith('.json')) files.push(fullPath);
  }
  return files;
}

function parseForFormat(format, graph) {
  if (format === 'docx') return parseDocx(graph);
  if (format === 'xlsx') return parseXlsx(graph);
  if (format === 'pptx') return parsePptx(graph);
  throw new Error(`Unsupported fixture format: ${format}`);
}

function representativeFixturePath(format, id) {
  const slug = id.replace(`${format}-`, '');
  return path.join(root, 'fixtures', format, 'representative', `${slug}.${format}`);
}

const manifestFiles = await walk(manifestsRoot);
const manifestEntries = await Promise.all(manifestFiles.map(async (filePath) => ({
  filePath,
  manifest: JSON.parse(await readFile(filePath, 'utf8'))
})));

const results = [];
for (const { manifest } of manifestEntries) {
  const fixturePath = representativeFixturePath(manifest.format, manifest.id);
  try {
    const bytes = new Uint8Array(await readFile(fixturePath));
    const graph = await openPackage(bytes);
    const parsed = parseForFormat(manifest.format, graph);
    const reopened = parseForFormat(manifest.format, await openPackage(serializeOfficeDocument(parsed)));
    results.push({
      id: manifest.id,
      format: manifest.format,
      fixturePath: path.relative(root, fixturePath).replaceAll('\\', '/'),
      parserOpen: true,
      parserRoundTrip: Boolean(reopened),
      officeStatus: manifest.reopenExpectations?.office ?? 'unknown',
      libreOfficeStatus: manifest.reopenExpectations?.libreoffice ?? 'unknown'
    });
  } catch (error) {
    results.push({
      id: manifest.id,
      format: manifest.format,
      fixturePath: path.relative(root, fixturePath).replaceAll('\\', '/'),
      parserOpen: false,
      parserRoundTrip: false,
      officeStatus: manifest.reopenExpectations?.office ?? 'unknown',
      libreOfficeStatus: manifest.reopenExpectations?.libreoffice ?? 'unknown',
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

await mkdir(path.dirname(resultsOutputPath), { recursive: true });
await writeFile(resultsOutputPath, JSON.stringify({ generatedAt: new Date().toISOString(), results }, null, 2));
console.log(JSON.stringify({ resultsOutputPath: path.relative(root, resultsOutputPath).replaceAll('\\', '/'), count: results.length }, null, 2));
