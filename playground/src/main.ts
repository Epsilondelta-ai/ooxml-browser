import { createBrowserSession } from '@ooxml/browser';
import { createOfficeEditor, replaceDocxParagraphText, setPresentationNotesText, setPresentationShapeText, setWorkbookCellValue, type EditableOfficeDocument, type OfficeEditor } from '@ooxml/editor';
import { renderOfficeDocumentToHtml } from '@ooxml/render';

const app = document.getElementById('app');

if (!app) {
  throw new Error('Missing #app mount point.');
}

app.innerHTML = `
  <section style="font-family: system-ui, sans-serif; max-width: 1200px; margin: 0 auto; padding: 24px; display: grid; gap: 16px;">
    <header style="display: grid; gap: 6px;">
      <h1 style="margin: 0;">OOXML Playground</h1>
      <p style="margin: 0; color: #475569;">Load an OOXML document, inspect its package summary, apply a small edit, and save the modified file.</p>
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
    currentEditor = createOfficeEditor(session.document);
    summary.textContent = JSON.stringify({
      packageSummary: session.packageSummary,
      documentSummary: session.documentSummary
    }, null, 2);
    renderPreview();
    renderEditorControls();
    saveButton.disabled = false;
    status.textContent = `Loaded ${file.name}`;
  } catch (error) {
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

  const bytes = currentEditor.serialize();
  const blob = new Blob([Uint8Array.from(bytes)], { type: 'application/octet-stream' });
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
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First paragraph</div>
          <input id="docx-input" value="${escapeHtml(firstParagraph)}" style="width: 100%; padding: 8px;" />
        </label>
      `;
      const input = window.document.getElementById('docx-input') as HTMLInputElement;
      input.addEventListener('input', () => {
        if (!currentEditor || currentEditor.document.kind !== 'docx') {
          return;
        }
        replaceDocxParagraphText(currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'docx' }>>, 0, 0, input.value);
        renderPreview();
      });
      break;
    }
    case 'xlsx': {
      const firstCell = officeDocument.sheets[0]?.rows[0]?.cells[0]?.value ?? '';
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Sheet1!A1</div>
          <input id="xlsx-input" value="${escapeHtml(firstCell)}" style="width: 100%; padding: 8px;" />
        </label>
      `;
      const input = window.document.getElementById('xlsx-input') as HTMLInputElement;
      input.addEventListener('input', () => {
        if (!currentEditor || currentEditor.document.kind !== 'xlsx') {
          return;
        }
        setWorkbookCellValue(currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'xlsx' }>>, 'Sheet1', 'A1', input.value);
        renderPreview();
      });
      break;
    }
    case 'pptx': {
      const shapeText = officeDocument.slides[0]?.shapes[0]?.text ?? '';
      const notesText = officeDocument.slides[0]?.notesText ?? '';
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First shape text</div>
          <input id="pptx-shape-input" value="${escapeHtml(shapeText)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Notes</div>
          <textarea id="pptx-notes-input" style="width: 100%; min-height: 96px; padding: 8px;">${escapeHtml(notesText)}</textarea>
        </label>
      `;
      const shapeInput = window.document.getElementById('pptx-shape-input') as HTMLInputElement;
      const notesInput = window.document.getElementById('pptx-notes-input') as HTMLTextAreaElement;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'pptx') {
          return;
        }
        setPresentationShapeText(currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'pptx' }>>, 0, 0, shapeInput.value);
        setPresentationNotesText(currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'pptx' }>>, 0, notesInput.value);
        renderPreview();
      };
      shapeInput.addEventListener('input', update);
      notesInput.addEventListener('input', update);
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
