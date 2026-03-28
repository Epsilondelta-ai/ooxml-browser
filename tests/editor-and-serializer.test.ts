import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { createOfficeEditor, replaceDocxParagraphText, setDocxCommentText, setPresentationNotesText, setPresentationShapeText, setWorkbookCellValue } from '@ooxml/editor';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx } from '@ooxml/xlsx';

import { createDocxFixture, createPptxFixture, createXlsxFixture } from './fixture-builders';

describe('editor transactions', () => {
  it('updates docx paragraphs and supports undo/redo', async () => {
    const graph = await openPackage(createDocxFixture());
    const editor = createOfficeEditor(parseDocx(graph));

    replaceDocxParagraphText(editor, 0, 0, 'Edited paragraph');
    expect(editor.document.stories[0]?.paragraphs[0]?.text).toBe('Edited paragraph');

    editor.undo();
    expect(editor.document.stories[0]?.paragraphs[0]?.text).toBe('Hello OOXML');

    editor.redo();
    expect(editor.document.stories[0]?.paragraphs[0]?.text).toBe('Edited paragraph');
  });

  it('updates workbook cells and presentation notes', async () => {
    const workbookEditor = createOfficeEditor(parseXlsx(await openPackage(createXlsxFixture())));
    setWorkbookCellValue(workbookEditor, 'Sheet1', 'A1', 'Updated cell');
    expect(workbookEditor.document.sheets[0]?.rows[0]?.cells[0]?.value).toBe('Updated cell');

    const presentationEditor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    setPresentationShapeText(presentationEditor, 0, 0, 'Updated slide');
    setPresentationNotesText(presentationEditor, 0, 'Updated note');
    expect(presentationEditor.document.slides[0]?.shapes[0]?.text).toBe('Updated slide');
    expect(presentationEditor.document.slides[0]?.notesText).toBe('Updated note');
  });
});

describe('serializer round-trips', () => {
  it('round-trips edited docx content', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createDocxFixture())));
    replaceDocxParagraphText(editor, 0, 0, 'Saved paragraph');

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.stories[0]?.paragraphs[0]?.text).toBe('Saved paragraph');
  });


  it('round-trips edited docx comments while preserving comment authors', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createDocxFixture())));
    setDocxCommentText(editor, '0', 'Updated comment');

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.comments[0]).toEqual({ id: '0', author: 'Codex', text: 'Updated comment' });
  });

  it('round-trips edited xlsx content', async () => {
    const editor = createOfficeEditor(parseXlsx(await openPackage(createXlsxFixture())));
    setWorkbookCellValue(editor, 'Sheet1', 'A1', 'Persisted');

    const reopened = parseXlsx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.sheets[0]?.rows[0]?.cells[0]?.value).toBe('Persisted');
  });



  it('round-trips notes edits when the notes relationship targets a non-canonical part name', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture({ notesTarget: '../notesSlides/customNotes.xml' }))));
    setPresentationNotesText(editor, 0, 'Custom target note');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.notesUri).toBe('/ppt/notesSlides/customNotes.xml');
    expect(reopened.slides[0]?.notesText).toBe('Custom target note');
  });

  it('does not persist notes edits when the source package has no notes part', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture({ withNotes: false }))));
    setPresentationNotesText(editor, 0, 'Ignored note');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.notesText).toBe('');
  });

  it('round-trips edited pptx content', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    setPresentationShapeText(editor, 0, 0, 'Persisted slide');
    setPresentationNotesText(editor, 0, 'Persisted notes');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.shapes[0]?.text).toBe('Persisted slide');
    expect(reopened.slides[0]?.notesText).toBe('Persisted notes');
  });
});
