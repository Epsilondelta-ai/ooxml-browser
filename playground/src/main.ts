import { createBrowserSession } from '@ooxml/browser';
import {
  replaceDocxParagraphText,
  setDocxParagraphStyle,
  setPresentationNotesText,
  setPresentationShapeText,
  setPresentationTransition,
  setWorkbookCellFormula,
  setWorkbookCellValue,
  setWorkbookSheetName,
  type EditableOfficeDocument,
  type OfficeEditor
} from '@ooxml/editor';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

const app = document.getElementById('app');

if (!app) {
  throw new Error('Missing #app mount point.');
}

app.innerHTML = `
  <section style="font-family: system-ui, sans-serif; max-width: 1200px; margin: 0 auto; padding: 24px; display: grid; gap: 16px;">
    <header style="display: grid; gap: 6px;">
      <h1 style="margin: 0;">OOXML Playground</h1>
      <p style="margin: 0; color: #475569;">Load an OOXML document, inspect its package summary, apply text and metadata edits, and save the modified file.</p>
    </header>

    <div style="display: flex; gap: 12px; flex-wrap: wrap; align-items: center;">
      <input id="file-input" type="file" accept=".docx,.xlsx,.pptx" />
      <button id="save-button" disabled>Save current document</button>
      <span id="status" style="color: #334155;">Waiting for a file…</span>
    </div>

    <div style="display: grid; grid-template-columns: minmax(280px, 360px) 1fr; gap: 16px; align-items: start;">
      <div style="display: grid; gap: 16px;">
        <section style="background: #0f172a; color: #e2e8f0; border-radius: 16px; padding: 16px;">
          <h2 style="margin-top: 0;">Summaries</h2>
          <pre id="summary" style="margin: 0; white-space: pre-wrap;"></pre>
        </section>
        <section style="background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 16px; padding: 16px;">
          <h2 style="margin-top: 0;">Quick edit</h2>
          <div id="editor-controls" style="display: grid; gap: 12px;"></div>
        </section>
      </div>
      <section style="background: white; border: 1px solid #cbd5e1; border-radius: 16px; padding: 20px; min-height: 320px;">
        <h2 style="margin-top: 0;">Preview</h2>
        <div id="preview"></div>
      </section>
    </div>
  </section>
`;

const fileInput = document.getElementById('file-input') as HTMLInputElement;
const saveButton = document.getElementById('save-button') as HTMLButtonElement;
const status = document.getElementById('status') as HTMLSpanElement;
const summary = document.getElementById('summary') as HTMLPreElement;
const preview = document.getElementById('preview') as HTMLDivElement;
const editorControls = document.getElementById('editor-controls') as HTMLDivElement;

let currentSession: Awaited<ReturnType<typeof createBrowserSession>> | null = null;
let currentEditor: OfficeEditor<EditableOfficeDocument> | null = null;
let currentFileName = 'document.ooxml';

fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) {
    return;
  }

  currentFileName = file.name;
  status.textContent = `Opening ${file.name}…`;

  try {
    const session = await createBrowserSession(file);
    currentSession = session;
    currentEditor = session.createEditor();
    summary.textContent = JSON.stringify({
      packageSummary: session.packageSummary,
      documentSummary: session.documentSummary
    }, null, 2);
    renderPreview();
    renderEditorControls();
    saveButton.disabled = false;
    status.textContent = `Loaded ${file.name}`;
  } catch (error) {
    currentSession = null;
    currentEditor = null;
    summary.textContent = '';
    preview.innerHTML = '';
    editorControls.innerHTML = '';
    saveButton.disabled = true;
    status.textContent = error instanceof Error ? error.message : 'Failed to open document.';
  }
});

saveButton.addEventListener('click', () => {
  if (!currentEditor) {
    return;
  }

  if (!currentSession) {
    return;
  }

  const blob = currentSession.save();
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = currentFileName;
  anchor.click();
  URL.revokeObjectURL(url);
  status.textContent = `Saved ${currentFileName}`;
});

function renderPreview(): void {
  if (!currentEditor) {
    preview.innerHTML = '';
    return;
  }

  preview.innerHTML = renderOfficeDocumentToHtml(currentEditor.document as never);
}

function renderEditorControls(): void {
  if (!currentEditor) {
    editorControls.innerHTML = '';
    return;
  }

  const officeDocument = currentEditor.document;
  switch (officeDocument.kind) {
    case 'docx': {
      const firstParagraph = officeDocument.stories[0]?.paragraphs[0]?.text ?? '';
      const firstStyle = officeDocument.stories[0]?.paragraphs[0]?.styleId ?? '';
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First paragraph</div>
          <input id="docx-input" value="${escapeHtml(firstParagraph)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Paragraph style</div>
          <input id="docx-style-input" value="${escapeHtml(firstStyle)}" style="width: 100%; padding: 8px;" />
        </label>
      `;
      const input = window.document.getElementById('docx-input') as HTMLInputElement;
      const styleInput = window.document.getElementById('docx-style-input') as HTMLInputElement;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'docx') {
          return;
        }
        replaceDocxParagraphText(currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'docx' }>>, 0, 0, input.value);
        setDocxParagraphStyle(
          currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'docx' }>>,
          'document',
          0,
          0,
          styleInput.value || undefined
        );
        renderPreview();
      };
      input.addEventListener('input', update);
      styleInput.addEventListener('input', update);
      break;
    }
    case 'xlsx': {
      const firstSheet = officeDocument.sheets[0];
      const firstCell = firstSheet?.rows[0]?.cells[0]?.value ?? '';
      const firstFormulaCell = firstSheet?.rows.flatMap((row) => row.cells).find((cell) => cell.formula);
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Sheet1!A1</div>
          <input id="xlsx-input" value="${escapeHtml(firstCell)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First sheet name</div>
          <input id="xlsx-sheet-input" value="${escapeHtml(firstSheet?.name ?? 'Sheet1')}" style="width: 100%; padding: 8px;" />
        </label>
        ${firstFormulaCell ? `<label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">${escapeHtml(firstFormulaCell.reference)} formula</div>
          <input id="xlsx-formula-input" value="${escapeHtml(firstFormulaCell.formula ?? '')}" style="width: 100%; padding: 8px;" />
        </label>` : ''}
      `;
      const input = window.document.getElementById('xlsx-input') as HTMLInputElement;
      const sheetInput = window.document.getElementById('xlsx-sheet-input') as HTMLInputElement;
      const formulaInput = firstFormulaCell ? window.document.getElementById('xlsx-formula-input') as HTMLInputElement : null;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'xlsx') {
          return;
        }
        const workbookEditor = currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'xlsx' }>>;
        const currentSheetName = workbookEditor.document.sheets[0]?.name ?? 'Sheet1';
        setWorkbookCellValue(workbookEditor, currentSheetName, 'A1', input.value);
        if (sheetInput.value && sheetInput.value !== currentSheetName) {
          setWorkbookSheetName(workbookEditor, currentSheetName, sheetInput.value);
        }
        if (formulaInput && firstFormulaCell) {
          const formulaSheetName = workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1';
          setWorkbookCellFormula(
            workbookEditor,
            formulaSheetName,
            firstFormulaCell.reference,
            formulaInput.value,
            firstFormulaCell.value
          );
        }
        renderPreview();
      };
      input.addEventListener('input', update);
      sheetInput.addEventListener('input', update);
      formulaInput?.addEventListener('input', update);
      break;
    }
    case 'pptx': {
      const shapeText = officeDocument.slides[0]?.shapes[0]?.text ?? '';
      const notesText = officeDocument.slides[0]?.notesText ?? '';
      const transitionType = officeDocument.slides[0]?.transition?.type ?? '';
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First shape text</div>
          <input id="pptx-shape-input" value="${escapeHtml(shapeText)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Notes</div>
          <textarea id="pptx-notes-input" style="width: 100%; min-height: 96px; padding: 8px;">${escapeHtml(notesText)}</textarea>
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Transition type</div>
          <input id="pptx-transition-input" value="${escapeHtml(transitionType)}" style="width: 100%; padding: 8px;" />
        </label>
      `;
      const shapeInput = window.document.getElementById('pptx-shape-input') as HTMLInputElement;
      const notesInput = window.document.getElementById('pptx-notes-input') as HTMLTextAreaElement;
      const transitionInput = window.document.getElementById('pptx-transition-input') as HTMLInputElement;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'pptx') {
          return;
        }
        const presentationEditor = currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'pptx' }>>;
        setPresentationShapeText(presentationEditor, 0, 0, shapeInput.value);
        setPresentationNotesText(presentationEditor, 0, notesInput.value);
        setPresentationTransition(presentationEditor, 0, transitionInput.value ? { type: transitionInput.value } : undefined);
        renderPreview();
      };
      shapeInput.addEventListener('input', update);
      notesInput.addEventListener('input', update);
      transitionInput.addEventListener('input', update);
      break;
    }
  }
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
