import { createBrowserSession } from '@ooxml/browser';
import {
  addDocxComment,
  addPresentationComment,
  replaceDocxParagraphText,
  setDocxCommentText,
  setDocxParagraphStyle,
  setPresentationCommentText,
  setPresentationNotesText,
  setPresentationShapeText,
  setPresentationSize,
  setPresentationSlideLayout,
  setPresentationSlideMaster,
  setPresentationSlideTheme,
  setPresentationTimingNodes,
  setPresentationTransition,
  setWorkbookCellFormula,
  setWorkbookCellStyle,
  setWorkbookCellValue,
  setWorkbookSheetName,
  setWorksheetChartTarget,
  setWorksheetMediaTarget,
  setWorksheetPageMargins,
  setWorksheetPageSetup,
  setWorksheetPrintArea,
  setWorksheetPrintTitles,
  setWorksheetSelection,
  upsertWorksheetComment,
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
      const firstComment = officeDocument.comments[0]?.text ?? '';
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First paragraph</div>
          <input id="docx-input" value="${escapeHtml(firstParagraph)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Paragraph style</div>
          <input id="docx-style-input" value="${escapeHtml(firstStyle)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First comment</div>
          <input id="docx-comment-input" value="${escapeHtml(firstComment)}" style="width: 100%; padding: 8px;" />
        </label>
      `;
      const input = window.document.getElementById('docx-input') as HTMLInputElement;
      const styleInput = window.document.getElementById('docx-style-input') as HTMLInputElement;
      const commentInput = window.document.getElementById('docx-comment-input') as HTMLInputElement;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'docx') {
          return;
        }
        const docxEditor = currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'docx' }>>;
        replaceDocxParagraphText(docxEditor, 0, 0, input.value);
        setDocxParagraphStyle(
          docxEditor,
          'document',
          0,
          0,
          styleInput.value || undefined
        );
        if (commentInput.value) {
          if (docxEditor.document.comments[0]) {
            setDocxCommentText(docxEditor, docxEditor.document.comments[0].id, commentInput.value);
          } else {
            addDocxComment(docxEditor, { id: 'playground-comment', author: 'Playground', text: commentInput.value });
          }
        }
        renderPreview();
      };
      input.addEventListener('input', update);
      styleInput.addEventListener('input', update);
      commentInput.addEventListener('input', update);
      break;
    }
    case 'xlsx': {
      const firstSheet = officeDocument.sheets[0];
      const firstCell = firstSheet?.rows[0]?.cells[0]?.value ?? '';
      const firstCellStyle = firstSheet?.rows[0]?.cells[0]?.styleIndex ?? '';
      const firstFormulaCell = firstSheet?.rows.flatMap((row) => row.cells).find((cell) => cell.formula);
      const commentB2 = firstSheet?.comments.find((comment) => comment.reference === 'B2');
      const printArea = officeDocument.definedNames.find((entry) => entry.name === '_xlnm.Print_Area' && entry.scopeSheetId === 0)?.reference.split('!')[1] ?? '';
      const printTitles = officeDocument.definedNames.find((entry) => entry.name === '_xlnm.Print_Titles' && entry.scopeSheetId === 0)?.reference ?? '';
      const selection = firstSheet?.selection;
      const pageOrientation = firstSheet?.pageSetup?.orientation ?? '';
      const topMargin = firstSheet?.pageMargins?.top ?? '';
      const firstChartTarget = firstSheet?.charts[0]?.targetUri ?? '';
      const firstMediaTarget = firstSheet?.media[0]?.targetUri ?? '';
      editorControls.innerHTML = `
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Sheet1!A1</div>
          <input id="xlsx-input" value="${escapeHtml(firstCell)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Sheet1!A1 style index</div>
          <input id="xlsx-style-input" value="${escapeHtml(String(firstCellStyle))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First sheet name</div>
          <input id="xlsx-sheet-input" value="${escapeHtml(firstSheet?.name ?? 'Sheet1')}" style="width: 100%; padding: 8px;" />
        </label>
        ${firstFormulaCell ? `<label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">${escapeHtml(firstFormulaCell.reference)} formula</div>
          <input id="xlsx-formula-input" value="${escapeHtml(firstFormulaCell.formula ?? '')}" style="width: 100%; padding: 8px;" />
        </label>` : ''}
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Print area</div>
          <input id="xlsx-print-area-input" value="${escapeHtml(printArea)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Print titles</div>
          <input id="xlsx-print-titles-input" value="${escapeHtml(printTitles)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Selection sqref</div>
          <input id="xlsx-selection-input" value="${escapeHtml(selection?.sqref ?? selection?.activeCell ?? '')}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Page orientation</div>
          <input id="xlsx-orientation-input" value="${escapeHtml(pageOrientation)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Top margin</div>
          <input id="xlsx-top-margin-input" value="${escapeHtml(String(topMargin))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart target</div>
          <input id="xlsx-chart-target-input" value="${escapeHtml(firstChartTarget)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First image target</div>
          <input id="xlsx-media-target-input" value="${escapeHtml(firstMediaTarget)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Sheet1!B2 comment</div>
          <input id="xlsx-comment-input" value="${escapeHtml(commentB2?.text ?? '')}" style="width: 100%; padding: 8px;" />
        </label>
      `;
      const input = window.document.getElementById('xlsx-input') as HTMLInputElement;
      const styleInput = window.document.getElementById('xlsx-style-input') as HTMLInputElement;
      const sheetInput = window.document.getElementById('xlsx-sheet-input') as HTMLInputElement;
      const formulaInput = firstFormulaCell ? window.document.getElementById('xlsx-formula-input') as HTMLInputElement : null;
      const printAreaInput = window.document.getElementById('xlsx-print-area-input') as HTMLInputElement;
      const printTitlesInput = window.document.getElementById('xlsx-print-titles-input') as HTMLInputElement;
      const selectionInput = window.document.getElementById('xlsx-selection-input') as HTMLInputElement;
      const orientationInput = window.document.getElementById('xlsx-orientation-input') as HTMLInputElement;
      const topMarginInput = window.document.getElementById('xlsx-top-margin-input') as HTMLInputElement;
      const chartTargetInput = window.document.getElementById('xlsx-chart-target-input') as HTMLInputElement;
      const mediaTargetInput = window.document.getElementById('xlsx-media-target-input') as HTMLInputElement;
      const commentInput = window.document.getElementById('xlsx-comment-input') as HTMLInputElement;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'xlsx') {
          return;
        }
        const workbookEditor = currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'xlsx' }>>;
        const currentSheetName = workbookEditor.document.sheets[0]?.name ?? 'Sheet1';
        setWorkbookCellValue(workbookEditor, currentSheetName, 'A1', input.value);
        if (styleInput.value !== '') {
          const styleIndex = Number(styleInput.value);
          if (!Number.isNaN(styleIndex)) {
            setWorkbookCellStyle(workbookEditor, currentSheetName, 'A1', styleIndex);
          }
        }
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
        if (printAreaInput.value) {
          setWorksheetPrintArea(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            printAreaInput.value
          );
        }
        if (printTitlesInput.value) {
          const [rows, columns] = printTitlesInput.value
            .split(',')
            .map((segment) => segment.split('!')[1]?.trim())
            .filter((value): value is string => Boolean(value));
          setWorksheetPrintTitles(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            { rows, columns }
          );
        }
        if (selectionInput.value) {
          setWorksheetSelection(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            { activeCell: selectionInput.value.split(':')[0], sqref: selectionInput.value }
          );
        }
        if (orientationInput.value) {
          setWorksheetPageSetup(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            {
              ...workbookEditor.document.sheets[0]?.pageSetup,
              orientation: orientationInput.value
            }
          );
        }
        if (topMarginInput.value !== '') {
          const topMarginValue = Number(topMarginInput.value);
          if (!Number.isNaN(topMarginValue)) {
            setWorksheetPageMargins(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              {
                ...workbookEditor.document.sheets[0]?.pageMargins,
                top: topMarginValue
              }
            );
          }
        }
        if (chartTargetInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartTarget(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartTargetInput.value
          );
        }
        if (mediaTargetInput.value && workbookEditor.document.sheets[0]?.media[0]) {
          setWorksheetMediaTarget(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            mediaTargetInput.value
          );
        }
        if (commentInput.value) {
          upsertWorksheetComment(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            'B2',
            commentInput.value,
            'Playground'
          );
        }
        renderPreview();
      };
      input.addEventListener('input', update);
      styleInput.addEventListener('input', update);
      sheetInput.addEventListener('input', update);
      formulaInput?.addEventListener('input', update);
      printAreaInput.addEventListener('input', update);
      printTitlesInput.addEventListener('input', update);
      selectionInput.addEventListener('input', update);
      orientationInput.addEventListener('input', update);
      topMarginInput.addEventListener('input', update);
      chartTargetInput.addEventListener('input', update);
      mediaTargetInput.addEventListener('input', update);
      commentInput.addEventListener('input', update);
      break;
    }
    case 'pptx': {
      const shapeText = officeDocument.slides[0]?.shapes[0]?.text ?? '';
      const notesText = officeDocument.slides[0]?.notesText ?? '';
      const commentText = officeDocument.slides[0]?.comments[0]?.text ?? '';
      const transitionType = officeDocument.slides[0]?.transition?.type ?? '';
      const firstTimingDuration = officeDocument.slides[0]?.timing?.nodes[0]?.duration ?? '';
      const layoutUri = officeDocument.slides[0]?.layoutUri ?? '';
      const masterUri = officeDocument.slides[0]?.masterUri ?? '';
      const themeUri = officeDocument.slides[0]?.themeUri ?? '';
      const sizeCx = officeDocument.size.cx;
      const sizeCy = officeDocument.size.cy;
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
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First comment</div>
          <input id="pptx-comment-input" value="${escapeHtml(commentText)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Transition type</div>
          <input id="pptx-transition-input" value="${escapeHtml(transitionType)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing duration</div>
          <input id="pptx-timing-duration-input" value="${escapeHtml(firstTimingDuration)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Layout URI</div>
          <input id="pptx-layout-input" value="${escapeHtml(layoutUri)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Master URI</div>
          <input id="pptx-master-input" value="${escapeHtml(masterUri)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Theme URI</div>
          <input id="pptx-theme-input" value="${escapeHtml(themeUri)}" style="width: 100%; padding: 8px;" />
        </label>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
          <label>
            <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Deck width (cx)</div>
            <input id="pptx-size-cx-input" value="${sizeCx}" style="width: 100%; padding: 8px;" />
          </label>
          <label>
            <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Deck height (cy)</div>
            <input id="pptx-size-cy-input" value="${sizeCy}" style="width: 100%; padding: 8px;" />
          </label>
        </div>
      `;
      const shapeInput = window.document.getElementById('pptx-shape-input') as HTMLInputElement;
      const notesInput = window.document.getElementById('pptx-notes-input') as HTMLTextAreaElement;
      const commentInput = window.document.getElementById('pptx-comment-input') as HTMLInputElement;
      const transitionInput = window.document.getElementById('pptx-transition-input') as HTMLInputElement;
      const timingDurationInput = window.document.getElementById('pptx-timing-duration-input') as HTMLInputElement;
      const layoutInput = window.document.getElementById('pptx-layout-input') as HTMLInputElement;
      const masterInput = window.document.getElementById('pptx-master-input') as HTMLInputElement;
      const themeInput = window.document.getElementById('pptx-theme-input') as HTMLInputElement;
      const sizeCxInput = window.document.getElementById('pptx-size-cx-input') as HTMLInputElement;
      const sizeCyInput = window.document.getElementById('pptx-size-cy-input') as HTMLInputElement;
      const update = () => {
        if (!currentEditor || currentEditor.document.kind !== 'pptx') {
          return;
        }
        const presentationEditor = currentEditor as OfficeEditor<Extract<EditableOfficeDocument, { kind: 'pptx' }>>;
        setPresentationShapeText(presentationEditor, 0, 0, shapeInput.value);
        setPresentationNotesText(presentationEditor, 0, notesInput.value);
        if (commentInput.value) {
          if (presentationEditor.document.slides[0]?.comments[0]) {
            setPresentationCommentText(presentationEditor, 0, 0, commentInput.value);
          } else {
            addPresentationComment(presentationEditor, 0, { author: 'Playground', text: commentInput.value });
          }
        }
        setPresentationTransition(presentationEditor, 0, transitionInput.value ? { type: transitionInput.value } : undefined);
        if (timingDurationInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, duration: timingDurationInput.value } : node)
          );
        }
        if (layoutInput.value) {
          setPresentationSlideLayout(presentationEditor, 0, layoutInput.value);
        }
        if (masterInput.value) {
          setPresentationSlideMaster(presentationEditor, 0, masterInput.value);
        }
        if (themeInput.value) {
          setPresentationSlideTheme(presentationEditor, 0, themeInput.value);
        }
        const cx = Number(sizeCxInput.value);
        const cy = Number(sizeCyInput.value);
        if (!Number.isNaN(cx) && !Number.isNaN(cy)) {
          setPresentationSize(presentationEditor, { cx, cy });
        }
        renderPreview();
      };
      shapeInput.addEventListener('input', update);
      notesInput.addEventListener('input', update);
      commentInput.addEventListener('input', update);
      transitionInput.addEventListener('input', update);
      timingDurationInput.addEventListener('input', update);
      layoutInput.addEventListener('input', update);
      masterInput.addEventListener('input', update);
      themeInput.addEventListener('input', update);
      sizeCxInput.addEventListener('input', update);
      sizeCyInput.addEventListener('input', update);
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
