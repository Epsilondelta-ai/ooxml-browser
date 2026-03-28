import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import {
  createOfficeEditor,
  replaceDocxParagraphText,
  setDocxCommentText,
  setPresentationCommentText,
  setPresentationNotesText,
  setPresentationShapeText,
  setWorkbookCellValue,
  setWorksheetCommentText
} from '@ooxml/editor';
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

function describePartDiff(originalGraph, editedGraph) {
  const originalUris = new Set(Object.keys(originalGraph.parts));
  const editedUris = new Set(Object.keys(editedGraph.parts));
  const changedParts = [];
  const addedParts = [];
  const removedParts = [];

  for (const uri of [...editedUris].sort()) {
    if (!originalUris.has(uri)) {
      addedParts.push(uri);
      continue;
    }

    const left = originalGraph.parts[uri]?.text ?? originalGraph.parts[uri]?.data?.toString();
    const right = editedGraph.parts[uri]?.text ?? editedGraph.parts[uri]?.data?.toString();
    if (left !== right) changedParts.push(uri);
  }

  for (const uri of [...originalUris].sort()) {
    if (!editedUris.has(uri)) removedParts.push(uri);
  }

  const totalOriginalPartCount = originalUris.size;
  const preservedPartCount = totalOriginalPartCount - changedParts.length - removedParts.length;

  return {
    changedParts,
    addedParts,
    removedParts,
    changedPartCount: changedParts.length,
    preservedPartCount,
    totalOriginalPartCount
  };
}

function chooseSpreadsheetMutationCell(workbook) {
  const sheet = workbook.sheets[0];
  const cells = sheet?.rows.flatMap((row) => row.cells) ?? [];
  return cells.find((cell) => cell.type === 'n' && !cell.formula)
    ?? cells.find((cell) => !cell.formula && !Number.isNaN(Number(cell.value)))
    ?? cells.find((cell) => !cell.formula)
    ?? cells[0];
}

function mutateDocument(format, document) {
  const editor = createOfficeEditor(document);
  if (format === 'docx') {
    if (editor.document.comments[0]) {
      setDocxCommentText(editor, editor.document.comments[0].id, 'Mutated comment');
      return { document: editor.document, mutation: 'comment-text-edit' };
    }

    replaceDocxParagraphText(editor, 0, 0, 'Mutated paragraph');
    return { document: editor.document, mutation: 'paragraph-text-edit' };
  }

  if (format === 'xlsx') {
    const workbook = editor.document;
    if (workbook.sheets[0]?.comments[0]) {
      setWorksheetCommentText(editor, workbook.sheets[0].name, workbook.sheets[0].comments[0].reference, 'Mutated comment');
      return { document: editor.document, mutation: 'comment-text-edit' };
    }

    const targetCell = chooseSpreadsheetMutationCell(workbook);
    setWorkbookCellValue(
      editor,
      workbook.sheets[0]?.name ?? 'Sheet1',
      targetCell?.reference ?? 'A1',
      '99'
    );
    return { document: editor.document, mutation: 'cell-value-edit' };
  }

  if (format === 'pptx') {
    const slide = editor.document.slides[0];
    if (slide?.comments[0]) {
      setPresentationCommentText(editor, 0, 0, 'Mutated comment');
      return { document: editor.document, mutation: 'comment-text-edit' };
    }

    if (slide?.notesUri) {
      setPresentationNotesText(editor, 0, 'Mutated note');
      return { document: editor.document, mutation: 'notes-text-edit' };
    }

    setPresentationShapeText(editor, 0, 0, 'Mutated slide');
    return { document: editor.document, mutation: 'shape-text-edit' };
  }

  return { document: editor.document, mutation: 'no-op' };
}

const manifestFiles = await walk(manifestsRoot);
const manifestEntries = await Promise.all(
  manifestFiles.map(async (filePath) => ({
    filePath,
    manifest: JSON.parse(await readFile(filePath, 'utf8'))
  }))
);

const results = [];
for (const { manifest } of manifestEntries) {
  const fixturePath = representativeFixturePath(manifest.format, manifest.id);
  try {
    const bytes = new Uint8Array(await readFile(fixturePath));
    const originalGraph = await openPackage(bytes);
    const parsed = parseForFormat(manifest.format, originalGraph);
    const reopened = parseForFormat(
      manifest.format,
      await openPackage(serializeOfficeDocument(parsed))
    );
    const { document: mutated, mutation } = mutateDocument(manifest.format, parsed);
    const editedBytes = serializeOfficeDocument(mutated);
    const editedGraph = await openPackage(editedBytes);
    const editedReopened = parseForFormat(manifest.format, editedGraph);
    const partDiff = describePartDiff(originalGraph, editedGraph);

    results.push({
      id: manifest.id,
      format: manifest.format,
      fixturePath: path.relative(root, fixturePath).replaceAll('\\', '/'),
      mutation,
      parserOpen: true,
      parserRoundTrip: Boolean(reopened),
      editedRoundTrip: Boolean(editedReopened),
      ...partDiff,
      officeStatus: manifest.reopenExpectations?.office ?? 'unknown',
      libreOfficeStatus: manifest.reopenExpectations?.libreoffice ?? 'unknown'
    });
  } catch (error) {
    results.push({
      id: manifest.id,
      format: manifest.format,
      fixturePath: path.relative(root, fixturePath).replaceAll('\\', '/'),
      mutation: 'failed-before-mutation',
      parserOpen: false,
      parserRoundTrip: false,
      editedRoundTrip: false,
      changedParts: [],
      addedParts: [],
      removedParts: [],
      changedPartCount: 0,
      preservedPartCount: 0,
      totalOriginalPartCount: 0,
      officeStatus: manifest.reopenExpectations?.office ?? 'unknown',
      libreOfficeStatus: manifest.reopenExpectations?.libreoffice ?? 'unknown',
      error: error instanceof Error ? error.message : String(error)
    });
  }
}

await mkdir(path.dirname(resultsOutputPath), { recursive: true });
await writeFile(
  resultsOutputPath,
  JSON.stringify({ generatedAt: new Date().toISOString(), results }, null, 2)
);
console.log(
  JSON.stringify(
    {
      resultsOutputPath: path.relative(root, resultsOutputPath).replaceAll('\\', '/'),
      count: results.length
    },
    null,
    2
  )
);
