import { describe, expect, it } from 'vitest';

import { openPackage } from '@ooxml/core';
import { parseDocx } from '@ooxml/docx';
import { createOfficeEditor, replaceDocxParagraphText, replaceDocxStoryParagraphText, setDocxCommentText, setDocxParagraphNumbering, setDocxParagraphRunStyle, setDocxParagraphStyle, setDocxSectionLayout, setDocxSectionReferenceType, setDocxTableCellText, setPresentationNotesText, setPresentationShapeText, setWorkbookCellValue } from '@ooxml/editor';
import { parsePptx } from '@ooxml/pptx';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx } from '@ooxml/xlsx';

import { createDocxFixture, createNumberedDocxFixture, createPptxFixture, createSectionedDocxFixture, createStyledDocxFixture, createXlsxFixture } from './fixture-builders';

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

  it('updates docx header/footer stories through story-aware helpers', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createSectionedDocxFixture())));
    replaceDocxStoryParagraphText(editor, 'header', 0, 0, 'Edited header');
    replaceDocxStoryParagraphText(editor, 'footer', 0, 0, 'Edited footer');

    expect(editor.document.stories.find((story) => story.kind === 'header')?.paragraphs[0]?.text).toBe('Edited header');
    expect(editor.document.stories.find((story) => story.kind === 'footer')?.paragraphs[0]?.text).toBe('Edited footer');
  });


  it('updates docx section layout metadata through editor helpers', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createSectionedDocxFixture())));
    setDocxSectionLayout(editor, 0, {
      pageSize: { width: 12240, height: 15840 },
      pageMargins: { top: 720, right: 960, bottom: 720, left: 960 }
    });

    expect(editor.document.sections[0]?.pageSize).toEqual({ width: 12240, height: 15840 });
    expect(editor.document.sections[0]?.pageMargins).toEqual({ top: 720, right: 960, bottom: 720, left: 960 });
  });

  it('updates docx header/footer reference types through editor helpers', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createSectionedDocxFixture())));
    setDocxSectionReferenceType(editor, 0, 'header', 0, 'first');
    setDocxSectionReferenceType(editor, 0, 'footer', 0, 'even');

    expect(editor.document.sections[0]?.headerReferences[0]?.type).toBe('first');
    expect(editor.document.sections[0]?.footerReferences[0]?.type).toBe('even');
  });

  it('updates docx paragraph style and numbering through editor helpers', async () => {
    const styledEditor = createOfficeEditor(parseDocx(await openPackage(createStyledDocxFixture())));
    setDocxParagraphStyle(styledEditor, 'document', 0, 0, 'Base');
    expect(styledEditor.document.stories[0]?.paragraphs[0]?.styleId).toBe('Base');

    const numberedEditor = createOfficeEditor(parseDocx(await openPackage(createNumberedDocxFixture())));
    setDocxParagraphNumbering(numberedEditor, 'document', 0, 0, { numId: '7', level: 1 });
    expect(numberedEditor.document.stories[0]?.paragraphs[0]?.numbering).toEqual({ numId: '7', level: 1 });
  });


  it('updates docx run formatting through editor helpers', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createStyledDocxFixture())));
    setDocxParagraphRunStyle(editor, 'document', 0, 0, 0, { bold: true, italic: false });

    expect(editor.document.stories[0]?.paragraphs[0]?.runs[0]).toMatchObject({ text: 'Styled heading', bold: true, italic: false });
  });

  it('updates docx table cells through story-aware helpers', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createDocxFixture())));
    setDocxTableCellText(editor, 'document', 0, 0, 0, 1, 'Edited cell');

    expect(editor.document.stories[0]?.tables[0]?.rows[0]?.cells[1]?.text).toBe('Edited cell');
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


  it('round-trips edited docx section layout metadata', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createSectionedDocxFixture())));
    setDocxSectionLayout(editor, 0, {
      pageSize: { width: 12240, height: 15840 },
      pageMargins: { top: 720, right: 960, bottom: 720, left: 960 }
    });

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.sections[0]?.pageSize).toEqual({ width: 12240, height: 15840 });
    expect(reopened.sections[0]?.pageMargins).toEqual({ top: 720, right: 960, bottom: 720, left: 960 });
  });

  it('round-trips edited docx header/footer reference types', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createSectionedDocxFixture())));
    setDocxSectionReferenceType(editor, 0, 'header', 0, 'first');
    setDocxSectionReferenceType(editor, 0, 'footer', 0, 'even');

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.sections[0]?.headerReferences[0]?.type).toBe('first');
    expect(reopened.sections[0]?.footerReferences[0]?.type).toBe('even');
  });

  it('round-trips edited docx paragraph style and numbering metadata', async () => {
    const styledEditor = createOfficeEditor(parseDocx(await openPackage(createStyledDocxFixture())));
    setDocxParagraphStyle(styledEditor, 'document', 0, 0, 'Base');
    const reopenedStyled = parseDocx(await openPackage(serializeOfficeDocument(styledEditor.document)));
    expect(reopenedStyled.stories[0]?.paragraphs[0]?.styleId).toBe('Base');

    const numberedEditor = createOfficeEditor(parseDocx(await openPackage(createNumberedDocxFixture())));
    setDocxParagraphNumbering(numberedEditor, 'document', 0, 0, { numId: '7', level: 1 });
    const reopenedNumbered = parseDocx(await openPackage(serializeOfficeDocument(numberedEditor.document)));
    expect(reopenedNumbered.stories[0]?.paragraphs[0]?.numbering).toEqual({ numId: '7', level: 1 });
  });


  it('round-trips edited docx run formatting metadata', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createStyledDocxFixture())));
    setDocxParagraphRunStyle(editor, 'document', 0, 0, 0, { bold: true, italic: false });

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.stories[0]?.paragraphs[0]?.runs[0]).toMatchObject({ text: 'Styled heading', bold: true, italic: false });
  });

  it('round-trips edited docx table cell content', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createDocxFixture())));
    setDocxTableCellText(editor, 'document', 0, 0, 0, 1, 'Saved cell');

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.stories[0]?.tables[0]?.rows[0]?.cells[1]?.text).toBe('Saved cell');
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

  it('creates and persists notes edits when the source package has no notes part', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture({ withNotes: false }))));
    setPresentationNotesText(editor, 0, 'Created note');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.notesUri).toBe('/ppt/notesSlides/notesSlide1.xml');
    expect(reopened.slides[0]?.notesText).toBe('Created note');
  });

  it('round-trips edited pptx content', async () => {
    const editor = createOfficeEditor(parsePptx(await openPackage(createPptxFixture())));
    setPresentationShapeText(editor, 0, 0, 'Persisted slide');
    setPresentationNotesText(editor, 0, 'Persisted notes');

    const reopened = parsePptx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.slides[0]?.shapes[0]?.text).toBe('Persisted slide');
    expect(reopened.slides[0]?.notesText).toBe('Persisted notes');
  });

  it('round-trips edited docx header/footer story content', async () => {
    const editor = createOfficeEditor(parseDocx(await openPackage(createSectionedDocxFixture())));
    replaceDocxStoryParagraphText(editor, 'header', 0, 0, 'Saved header');
    replaceDocxStoryParagraphText(editor, 'footer', 0, 0, 'Saved footer');

    const reopened = parseDocx(await openPackage(serializeOfficeDocument(editor.document)));
    expect(reopened.stories.find((story) => story.kind === 'header')?.paragraphs[0]?.text).toBe('Saved header');
    expect(reopened.stories.find((story) => story.kind === 'footer')?.paragraphs[0]?.text).toBe('Saved footer');
  });
});
