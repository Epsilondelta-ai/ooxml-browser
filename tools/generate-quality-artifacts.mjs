import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

const root = process.cwd();
const manifestsRoot = path.join(root, 'fixtures', 'manifests');
const benchmarkReportPath = path.join(root, 'benchmarks', 'reports', 'latest-benchmark-results.json');
const matrixOutputPath = path.join(root, 'docs', 'quality', 'interop-matrix.md');
const benchmarkOutputPath = path.join(root, 'docs', 'quality', 'benchmark-baseline.md');

async function walk(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...await walk(fullPath));
    } else if (entry.isFile() && entry.name.endsWith('.json')) {
      files.push(fullPath);
    }
  }
  return files;
}

function readJson(filePath) {
  return readFile(filePath, 'utf8').then((text) => JSON.parse(text));
}

function relative(filePath) {
  return path.relative(root, filePath).replaceAll('\\', '/');
}

const manifestFiles = await walk(manifestsRoot);
const manifests = await Promise.all(manifestFiles.map(async (filePath) => ({
  filePath,
  relativePath: relative(filePath),
  manifest: await readJson(filePath)
})));

const grouped = manifests.reduce((acc, entry) => {
  const key = entry.manifest.format;
  acc[key] ??= [];
  acc[key].push(entry);
  return acc;
}, {});

const matrixSections = ['# Interoperability Matrix', '', 'Generated from `fixtures/manifests/**`.', ''];
for (const format of Object.keys(grouped).sort()) {
  matrixSections.push(`## ${format.toUpperCase()}`, '');
  matrixSections.push('| Fixture | Tags | Parser | Office | LibreOffice | Supported operations |');
  matrixSections.push('| --- | --- | --- | --- | --- | --- |');
  for (const entry of grouped[format].sort((left, right) => left.manifest.id.localeCompare(right.manifest.id))) {
    const { manifest } = entry;
    matrixSections.push(`| ${manifest.id} | ${manifest.featureTags.join(', ')} | ${String(manifest.reopenExpectations?.parser ?? '')} | ${String(manifest.reopenExpectations?.office ?? '')} | ${String(manifest.reopenExpectations?.libreoffice ?? '')} | ${manifest.supportedOperations.join(', ')} |`);
  }
  matrixSections.push('');
}

let benchmarkMarkdown = '# Benchmark Baseline\n\nLatest benchmark summary generated from `benchmarks/reports/latest-benchmark-results.json`.\n';
try {
  const benchmark = await readJson(benchmarkReportPath);
  benchmarkMarkdown += '\n| Label | Open ms | Parse ms | Render ms | Serialize ms | Output bytes | HTML length |\n';
  benchmarkMarkdown += '| --- | ---: | ---: | ---: | ---: | ---: | ---: |\n';
  for (const result of benchmark.results ?? []) {
    benchmarkMarkdown += `| ${result.label} | ${result.openMs} | ${result.parseMs} | ${result.renderMs} | ${result.serializeMs} | ${result.outputBytes} | ${result.htmlLength} |\n`;
  }
} catch {
  benchmarkMarkdown += '\nBenchmark report not found yet. Run `npm run bench` first.\n';
}

await mkdir(path.dirname(matrixOutputPath), { recursive: true });
await writeFile(matrixOutputPath, `${matrixSections.join('\n')}\n`);
await writeFile(benchmarkOutputPath, benchmarkMarkdown);
console.log(JSON.stringify({ matrixOutputPath: relative(matrixOutputPath), benchmarkOutputPath: relative(benchmarkOutputPath), manifestCount: manifests.length }, null, 2));
