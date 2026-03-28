import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import path from 'node:path';

const root = process.cwd();
const manifestsRoot = path.join(root, 'fixtures', 'manifests');
const attestationsRoot = path.join(root, 'fixtures', 'attestations');
const outputPath = path.join(root, 'benchmarks', 'reports', 'latest-attestation-report.json');

async function walk(dir) {
  const entries = await readdir(dir, { withFileTypes: true }).catch(() => []);
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) files.push(...await walk(fullPath));
    else if (entry.isFile() && entry.name.endsWith('.json')) files.push(fullPath);
  }
  return files;
}

function relative(filePath) {
  return path.relative(root, filePath).replaceAll('\\', '/');
}

function normalizeSide(entry, fallbackStatus) {
  if (!entry) {
    return { status: fallbackStatus, validatedAt: null, version: null, notes: null };
  }

  return {
    status: entry.status ?? fallbackStatus,
    validatedAt: entry.validatedAt ?? null,
    version: entry.version ?? null,
    notes: entry.notes ?? null
  };
}

const manifestFiles = await walk(manifestsRoot);
const manifests = await Promise.all(
  manifestFiles.map(async (filePath) => ({
    path: relative(filePath),
    manifest: JSON.parse(await readFile(filePath, 'utf8'))
  }))
);
const attestationFiles = await walk(attestationsRoot);
const attestations = new Map(
  await Promise.all(
    attestationFiles.map(async (filePath) => {
      const content = JSON.parse(await readFile(filePath, 'utf8'));
      return [content.fixtureId, { path: relative(filePath), content }];
    })
  )
);

const results = manifests
  .map(({ manifest, path: manifestPath }) => {
    const attestation = attestations.get(manifest.id);
    return {
      fixtureId: manifest.id,
      format: manifest.format,
      manifestPath,
      attestationPath: attestation?.path ?? null,
      office: normalizeSide(attestation?.content?.office, manifest.reopenExpectations?.office ?? 'pending'),
      libreOffice: normalizeSide(attestation?.content?.libreOffice, manifest.reopenExpectations?.libreoffice ?? 'pending')
    };
  })
  .sort((left, right) => left.fixtureId.localeCompare(right.fixtureId));

const summary = results.reduce(
  (acc, result) => {
    for (const side of ['office', 'libreOffice']) {
      const status = result[side].status;
      acc[side][status] = (acc[side][status] ?? 0) + 1;
    }
    return acc;
  },
  { office: {}, libreOffice: {} }
);

await mkdir(path.dirname(outputPath), { recursive: true });
await writeFile(outputPath, JSON.stringify({ generatedAt: new Date().toISOString(), results, summary }, null, 2));
console.log(JSON.stringify({ outputPath: relative(outputPath), count: results.length }, null, 2));
