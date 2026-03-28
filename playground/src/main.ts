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
  setWorksheetChartBubbleScale,
  setWorksheetChartCategoryAxisPosition,
  setWorksheetChartCategoryAxisTitle,
  setWorksheetChartDataLabels,
  setWorksheetChartDataLabelVisibility,
  setWorksheetChartDisplayBlanksAs,
  setWorksheetChartFirstSliceAngle,
  setWorksheetChartGapWidth,
  setWorksheetChartGrouping,
  setWorksheetChartHoleSize,
  setWorksheetChartLegendPosition,
  setWorksheetChartName,
  setWorksheetChartOverlap,
  setWorksheetChartPlotVisibleOnly,
  setWorksheetChartScatterStyle,
  setWorksheetChartSeriesExplosion,
  setWorksheetChartSeriesInvertIfNegative,
  setWorksheetChartSeriesMarker,
  setWorksheetChartSeriesName,
  setWorksheetChartShowNegativeBubbles,
  setWorksheetChartSizeRepresents,
  setWorksheetChartSmooth,
  setWorksheetChartTarget,
  setWorksheetChartTitle,
  setWorksheetChartType,
  setWorksheetChartValueAxisPosition,
  setWorksheetChartValueAxisTitle,
  setWorksheetChartVaryColors,
  setWorksheetMediaTarget,
  setWorksheetPageMargins,
  setWorksheetPageSetup,
  setWorksheetPrintArea,
  setWorksheetPrintTitles,
  setWorksheetSelection,
  setWorksheetThreadedCommentPerson,
  setWorksheetThreadedCommentText,
  upsertWorkbookThreadedCommentPerson,
  upsertWorksheetComment,
  upsertWorksheetThreadedComment,
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
      const firstChartName = firstSheet?.charts[0]?.name ?? '';
      const firstChartType = firstSheet?.charts[0]?.chartType ?? '';
      const firstChartScatterStyle = firstSheet?.charts[0]?.scatterStyle ?? '';
      const firstChartBubbleScale = firstSheet?.charts[0]?.bubbleScale ?? '';
      const firstChartShowNegativeBubbles = firstSheet?.charts[0]?.showNegativeBubbles;
      const firstChartSizeRepresents = firstSheet?.charts[0]?.sizeRepresents ?? '';
      const firstChartSmooth = firstSheet?.charts[0]?.smooth;
      const firstChartPlotVisibleOnly = firstSheet?.charts[0]?.plotVisibleOnly;
      const firstChartDisplayBlanksAs = firstSheet?.charts[0]?.displayBlanksAs ?? '';
      const firstChartGrouping = firstSheet?.charts[0]?.grouping ?? '';
      const firstChartOverlap = firstSheet?.charts[0]?.overlap ?? '';
      const firstChartVaryColors = firstSheet?.charts[0]?.varyColors;
      const firstChartGapWidth = firstSheet?.charts[0]?.gapWidth ?? '';
      const firstChartTarget = firstSheet?.charts[0]?.targetUri ?? '';
      const firstChartTitle = firstSheet?.charts[0]?.title ?? '';
      const firstChartFirstSlice = firstSheet?.charts[0]?.firstSliceAngle ?? '';
      const firstChartHoleSize = firstSheet?.charts[0]?.holeSize ?? '';
      const firstChartLegend = firstSheet?.charts[0]?.legendPosition ?? '';
      const firstChartCategoryAxisTitle = firstSheet?.charts[0]?.categoryAxisTitle ?? '';
      const firstChartCategoryAxisPosition = firstSheet?.charts[0]?.categoryAxisPosition ?? '';
      const firstChartValueAxisTitle = firstSheet?.charts[0]?.valueAxisTitle ?? '';
      const firstChartValueAxisPosition = firstSheet?.charts[0]?.valueAxisPosition ?? '';
      const firstChartDataLabelPosition = firstSheet?.charts[0]?.dataLabels?.position ?? '';
      const firstChartDataLabelSeparator = firstSheet?.charts[0]?.dataLabels?.separator ?? '';
      const firstChartDataLabelSeries = firstSheet?.charts[0]?.dataLabels?.showSeriesName;
      const firstChartDataLabelLegend = firstSheet?.charts[0]?.dataLabels?.showLegendKey;
      const firstChartDataLabelLeader = firstSheet?.charts[0]?.dataLabels?.showLeaderLines;
      const firstChartDataLabelPercent = firstSheet?.charts[0]?.dataLabels?.showPercent;
      const firstChartDataLabelBubble = firstSheet?.charts[0]?.dataLabels?.showBubbleSize;
      const firstChartSeriesName = firstSheet?.charts[0]?.seriesNames[0] ?? '';
      const firstChartSeriesInvert = firstSheet?.charts[0]?.series[0]?.invertIfNegative;
      const firstChartSeriesMarker = firstSheet?.charts[0]?.series[0]?.markerSymbol ?? '';
      const firstChartSeriesExplosion = firstSheet?.charts[0]?.series[0]?.explosion ?? '';
      const firstMediaTarget = firstSheet?.media[0]?.targetUri ?? '';
      const firstThreadedComment = firstSheet?.threadedComments[0];
      const firstThreadedCommentText = firstThreadedComment?.text ?? '';
      const firstThreadedCommentPerson = firstThreadedComment?.personId ?? '';
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
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart name</div>
          <input id="xlsx-chart-name-input" value="${escapeHtml(firstChartName)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart type</div>
          <input id="xlsx-chart-type-input" value="${escapeHtml(firstChartType)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Scatter style</div>
          <input id="xlsx-chart-scatter-style-input" value="${escapeHtml(firstChartScatterStyle)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Bubble scale</div>
          <input id="xlsx-chart-bubble-scale-input" value="${escapeHtml(String(firstChartBubbleScale))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Show negative bubbles</div>
          <input id="xlsx-chart-show-negative-bubbles-input" value="${escapeHtml(firstChartShowNegativeBubbles === undefined ? '' : String(firstChartShowNegativeBubbles))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Size represents</div>
          <input id="xlsx-chart-size-represents-input" value="${escapeHtml(firstChartSizeRepresents)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Smooth line</div>
          <input id="xlsx-chart-smooth-input" value="${escapeHtml(firstChartSmooth === undefined ? '' : String(firstChartSmooth))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Plot visible only</div>
          <input id="xlsx-chart-plot-visible-input" value="${escapeHtml(firstChartPlotVisibleOnly === undefined ? '' : String(firstChartPlotVisibleOnly))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Display blanks as</div>
          <input id="xlsx-chart-display-blanks-input" value="${escapeHtml(firstChartDisplayBlanksAs)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Chart grouping</div>
          <input id="xlsx-chart-grouping-input" value="${escapeHtml(firstChartGrouping)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Chart overlap</div>
          <input id="xlsx-chart-overlap-input" value="${escapeHtml(String(firstChartOverlap))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Vary colors</div>
          <input id="xlsx-chart-vary-colors-input" value="${escapeHtml(firstChartVaryColors === undefined ? '' : String(firstChartVaryColors))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Gap width</div>
          <input id="xlsx-chart-gap-width-input" value="${escapeHtml(String(firstChartGapWidth))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart target</div>
          <input id="xlsx-chart-target-input" value="${escapeHtml(firstChartTarget)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart title</div>
          <input id="xlsx-chart-title-input" value="${escapeHtml(firstChartTitle)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First slice angle</div>
          <input id="xlsx-chart-first-slice-input" value="${escapeHtml(String(firstChartFirstSlice))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Hole size</div>
          <input id="xlsx-chart-hole-size-input" value="${escapeHtml(String(firstChartHoleSize))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart legend</div>
          <input id="xlsx-chart-legend-input" value="${escapeHtml(firstChartLegend)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Category axis title</div>
          <input id="xlsx-chart-cat-axis-input" value="${escapeHtml(firstChartCategoryAxisTitle)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Category axis position</div>
          <input id="xlsx-chart-cat-axis-pos-input" value="${escapeHtml(firstChartCategoryAxisPosition)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Value axis title</div>
          <input id="xlsx-chart-val-axis-input" value="${escapeHtml(firstChartValueAxisTitle)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Value axis position</div>
          <input id="xlsx-chart-val-axis-pos-input" value="${escapeHtml(firstChartValueAxisPosition)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Data label position</div>
          <input id="xlsx-chart-dlbl-pos-input" value="${escapeHtml(firstChartDataLabelPosition)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Data label separator</div>
          <input id="xlsx-chart-dlbl-separator-input" value="${escapeHtml(firstChartDataLabelSeparator)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Show series name</div>
          <input id="xlsx-chart-dlbl-series-input" value="${escapeHtml(firstChartDataLabelSeries === undefined ? '' : String(firstChartDataLabelSeries))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Show legend key</div>
          <input id="xlsx-chart-dlbl-legend-input" value="${escapeHtml(firstChartDataLabelLegend === undefined ? '' : String(firstChartDataLabelLegend))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Show leader lines</div>
          <input id="xlsx-chart-dlbl-leader-input" value="${escapeHtml(firstChartDataLabelLeader === undefined ? '' : String(firstChartDataLabelLeader))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Show percent</div>
          <input id="xlsx-chart-dlbl-percent-input" value="${escapeHtml(firstChartDataLabelPercent === undefined ? '' : String(firstChartDataLabelPercent))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Show bubble size</div>
          <input id="xlsx-chart-dlbl-bubble-input" value="${escapeHtml(firstChartDataLabelBubble === undefined ? '' : String(firstChartDataLabelBubble))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First chart series</div>
          <input id="xlsx-chart-series-input" value="${escapeHtml(firstChartSeriesName)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Series invert negative</div>
          <input id="xlsx-chart-series-invert-input" value="${escapeHtml(firstChartSeriesInvert === undefined ? '' : String(firstChartSeriesInvert))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Series marker</div>
          <input id="xlsx-chart-series-marker-input" value="${escapeHtml(firstChartSeriesMarker)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Series explosion</div>
          <input id="xlsx-chart-series-explosion-input" value="${escapeHtml(String(firstChartSeriesExplosion))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First image target</div>
          <input id="xlsx-media-target-input" value="${escapeHtml(firstMediaTarget)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Sheet1!B2 comment</div>
          <input id="xlsx-comment-input" value="${escapeHtml(commentB2?.text ?? '')}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Threaded comment text</div>
          <input id="xlsx-threaded-comment-input" value="${escapeHtml(firstThreadedCommentText)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Threaded person ID</div>
          <input id="xlsx-threaded-person-input" value="${escapeHtml(firstThreadedCommentPerson)}" style="width: 100%; padding: 8px;" />
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
      const chartNameInput = window.document.getElementById('xlsx-chart-name-input') as HTMLInputElement;
      const chartTypeInput = window.document.getElementById('xlsx-chart-type-input') as HTMLInputElement;
      const chartScatterStyleInput = window.document.getElementById('xlsx-chart-scatter-style-input') as HTMLInputElement;
      const chartBubbleScaleInput = window.document.getElementById('xlsx-chart-bubble-scale-input') as HTMLInputElement;
      const chartShowNegativeBubblesInput = window.document.getElementById('xlsx-chart-show-negative-bubbles-input') as HTMLInputElement;
      const chartSizeRepresentsInput = window.document.getElementById('xlsx-chart-size-represents-input') as HTMLInputElement;
      const chartSmoothInput = window.document.getElementById('xlsx-chart-smooth-input') as HTMLInputElement;
      const chartPlotVisibleInput = window.document.getElementById('xlsx-chart-plot-visible-input') as HTMLInputElement;
      const chartDisplayBlanksInput = window.document.getElementById('xlsx-chart-display-blanks-input') as HTMLInputElement;
      const chartGroupingInput = window.document.getElementById('xlsx-chart-grouping-input') as HTMLInputElement;
      const chartOverlapInput = window.document.getElementById('xlsx-chart-overlap-input') as HTMLInputElement;
      const chartVaryColorsInput = window.document.getElementById('xlsx-chart-vary-colors-input') as HTMLInputElement;
      const chartGapWidthInput = window.document.getElementById('xlsx-chart-gap-width-input') as HTMLInputElement;
      const chartTargetInput = window.document.getElementById('xlsx-chart-target-input') as HTMLInputElement;
      const chartTitleInput = window.document.getElementById('xlsx-chart-title-input') as HTMLInputElement;
      const chartFirstSliceInput = window.document.getElementById('xlsx-chart-first-slice-input') as HTMLInputElement;
      const chartHoleSizeInput = window.document.getElementById('xlsx-chart-hole-size-input') as HTMLInputElement;
      const chartLegendInput = window.document.getElementById('xlsx-chart-legend-input') as HTMLInputElement;
      const chartCategoryAxisInput = window.document.getElementById('xlsx-chart-cat-axis-input') as HTMLInputElement;
      const chartCategoryAxisPositionInput = window.document.getElementById('xlsx-chart-cat-axis-pos-input') as HTMLInputElement;
      const chartValueAxisInput = window.document.getElementById('xlsx-chart-val-axis-input') as HTMLInputElement;
      const chartValueAxisPositionInput = window.document.getElementById('xlsx-chart-val-axis-pos-input') as HTMLInputElement;
      const chartDataLabelPositionInput = window.document.getElementById('xlsx-chart-dlbl-pos-input') as HTMLInputElement;
      const chartDataLabelSeparatorInput = window.document.getElementById('xlsx-chart-dlbl-separator-input') as HTMLInputElement;
      const chartDataLabelSeriesInput = window.document.getElementById('xlsx-chart-dlbl-series-input') as HTMLInputElement;
      const chartDataLabelLegendInput = window.document.getElementById('xlsx-chart-dlbl-legend-input') as HTMLInputElement;
      const chartDataLabelLeaderInput = window.document.getElementById('xlsx-chart-dlbl-leader-input') as HTMLInputElement;
      const chartDataLabelPercentInput = window.document.getElementById('xlsx-chart-dlbl-percent-input') as HTMLInputElement;
      const chartDataLabelBubbleInput = window.document.getElementById('xlsx-chart-dlbl-bubble-input') as HTMLInputElement;
      const chartSeriesInput = window.document.getElementById('xlsx-chart-series-input') as HTMLInputElement;
      const chartSeriesInvertInput = window.document.getElementById('xlsx-chart-series-invert-input') as HTMLInputElement;
      const chartSeriesMarkerInput = window.document.getElementById('xlsx-chart-series-marker-input') as HTMLInputElement;
      const chartSeriesExplosionInput = window.document.getElementById('xlsx-chart-series-explosion-input') as HTMLInputElement;
      const mediaTargetInput = window.document.getElementById('xlsx-media-target-input') as HTMLInputElement;
      const commentInput = window.document.getElementById('xlsx-comment-input') as HTMLInputElement;
      const threadedCommentInput = window.document.getElementById('xlsx-threaded-comment-input') as HTMLInputElement;
      const threadedPersonInput = window.document.getElementById('xlsx-threaded-person-input') as HTMLInputElement;
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
        if (chartNameInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartName(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartNameInput.value
          );
        }
        if (chartTypeInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartType(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartTypeInput.value
          );
        }
        if (chartScatterStyleInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartScatterStyle(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartScatterStyleInput.value
          );
        }
        if (chartBubbleScaleInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          const bubbleScale = Number(chartBubbleScaleInput.value);
          if (!Number.isNaN(bubbleScale)) {
            setWorksheetChartBubbleScale(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              0,
              bubbleScale
            );
          }
        }
        if (chartShowNegativeBubblesInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartShowNegativeBubbles(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartShowNegativeBubblesInput.value === 'true'
          );
        }
        if (chartSizeRepresentsInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartSizeRepresents(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartSizeRepresentsInput.value
          );
        }
        if (chartSmoothInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartSmooth(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartSmoothInput.value === 'true'
          );
        }
        if (chartPlotVisibleInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartPlotVisibleOnly(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartPlotVisibleInput.value === 'true'
          );
        }
        if (chartDisplayBlanksInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartDisplayBlanksAs(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartDisplayBlanksInput.value
          );
        }
        if (chartGroupingInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartGrouping(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartGroupingInput.value
          );
        }
        if (chartOverlapInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          const overlap = Number(chartOverlapInput.value);
          if (!Number.isNaN(overlap)) {
            setWorksheetChartOverlap(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              0,
              overlap
            );
          }
        }
        if (chartVaryColorsInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartVaryColors(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartVaryColorsInput.value === 'true'
          );
        }
        if (chartGapWidthInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          const gapWidth = Number(chartGapWidthInput.value);
          if (!Number.isNaN(gapWidth)) {
            setWorksheetChartGapWidth(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              0,
              gapWidth
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
        if (chartTitleInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartTitle(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartTitleInput.value
          );
        }
        if (chartFirstSliceInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          const angle = Number(chartFirstSliceInput.value);
          if (!Number.isNaN(angle)) {
            setWorksheetChartFirstSliceAngle(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              0,
              angle
            );
          }
        }
        if (chartHoleSizeInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          const holeSize = Number(chartHoleSizeInput.value);
          if (!Number.isNaN(holeSize)) {
            setWorksheetChartHoleSize(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              0,
              holeSize
            );
          }
        }
        if (chartLegendInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartLegendPosition(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartLegendInput.value
          );
        }
        if (chartCategoryAxisInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartCategoryAxisTitle(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartCategoryAxisInput.value
          );
        }
        if (chartCategoryAxisPositionInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartCategoryAxisPosition(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartCategoryAxisPositionInput.value
          );
        }
        if (chartValueAxisInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartValueAxisTitle(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartValueAxisInput.value
          );
        }
        if (chartValueAxisPositionInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartValueAxisPosition(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            chartValueAxisPositionInput.value
          );
        }
        if (chartDataLabelPositionInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartDataLabels(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            {
              ...workbookEditor.document.sheets[0]?.charts[0]?.dataLabels,
              position: chartDataLabelPositionInput.value
            }
          );
        }
        if (chartDataLabelSeparatorInput.value && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartDataLabels(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            {
              ...workbookEditor.document.sheets[0]?.charts[0]?.dataLabels,
              separator: chartDataLabelSeparatorInput.value
            }
          );
        }
        if ((chartDataLabelSeriesInput.value || chartDataLabelLegendInput.value || chartDataLabelLeaderInput.value || chartDataLabelPercentInput.value || chartDataLabelBubbleInput.value) && workbookEditor.document.sheets[0]?.charts[0]) {
          setWorksheetChartDataLabelVisibility(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            {
              showSeriesName: chartDataLabelSeriesInput.value ? chartDataLabelSeriesInput.value === 'true' : undefined,
              showLegendKey: chartDataLabelLegendInput.value ? chartDataLabelLegendInput.value === 'true' : undefined,
              showLeaderLines: chartDataLabelLeaderInput.value ? chartDataLabelLeaderInput.value === 'true' : undefined,
              showPercent: chartDataLabelPercentInput.value ? chartDataLabelPercentInput.value === 'true' : undefined,
              showBubbleSize: chartDataLabelBubbleInput.value ? chartDataLabelBubbleInput.value === 'true' : undefined
            }
          );
        }
        if (chartSeriesInput.value && workbookEditor.document.sheets[0]?.charts[0]?.seriesNames[0] !== undefined) {
          setWorksheetChartSeriesName(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            0,
            chartSeriesInput.value
          );
        }
        if (chartSeriesInvertInput.value && workbookEditor.document.sheets[0]?.charts[0]?.series[0]) {
          setWorksheetChartSeriesInvertIfNegative(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            0,
            chartSeriesInvertInput.value === 'true'
          );
        }
        if (chartSeriesMarkerInput.value && workbookEditor.document.sheets[0]?.charts[0]?.series[0]) {
          setWorksheetChartSeriesMarker(
            workbookEditor,
            workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
            0,
            0,
            {
              symbol: chartSeriesMarkerInput.value,
              size: workbookEditor.document.sheets[0]?.charts[0]?.series[0]?.markerSize
            }
          );
        }
        if (chartSeriesExplosionInput.value && workbookEditor.document.sheets[0]?.charts[0]?.series[0]) {
          const explosion = Number(chartSeriesExplosionInput.value);
          if (!Number.isNaN(explosion)) {
            setWorksheetChartSeriesExplosion(
              workbookEditor,
              workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1',
              0,
              0,
              explosion
            );
          }
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
        if (threadedPersonInput.value) {
          upsertWorkbookThreadedCommentPerson(workbookEditor, threadedPersonInput.value, threadedPersonInput.value);
        }
        if (threadedCommentInput.value && threadedPersonInput.value) {
          const sheetName = workbookEditor.document.sheets[0]?.name ?? sheetInput.value ?? 'Sheet1';
          if (workbookEditor.document.sheets[0]?.threadedComments[0]) {
            setWorksheetThreadedCommentText(workbookEditor, sheetName, workbookEditor.document.sheets[0].threadedComments[0].reference, threadedCommentInput.value);
            setWorksheetThreadedCommentPerson(workbookEditor, sheetName, workbookEditor.document.sheets[0].threadedComments[0].reference, threadedPersonInput.value);
          } else {
            upsertWorksheetThreadedComment(workbookEditor, sheetName, 'A1', threadedCommentInput.value, threadedPersonInput.value);
          }
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
      chartNameInput.addEventListener('input', update);
      chartTypeInput.addEventListener('input', update);
      chartScatterStyleInput.addEventListener('input', update);
      chartBubbleScaleInput.addEventListener('input', update);
      chartShowNegativeBubblesInput.addEventListener('input', update);
      chartSizeRepresentsInput.addEventListener('input', update);
      chartSmoothInput.addEventListener('input', update);
      chartPlotVisibleInput.addEventListener('input', update);
      chartDisplayBlanksInput.addEventListener('input', update);
      chartGroupingInput.addEventListener('input', update);
      chartOverlapInput.addEventListener('input', update);
      chartVaryColorsInput.addEventListener('input', update);
      chartGapWidthInput.addEventListener('input', update);
      chartTargetInput.addEventListener('input', update);
      chartTitleInput.addEventListener('input', update);
      chartFirstSliceInput.addEventListener('input', update);
      chartHoleSizeInput.addEventListener('input', update);
      chartLegendInput.addEventListener('input', update);
      chartCategoryAxisInput.addEventListener('input', update);
      chartCategoryAxisPositionInput.addEventListener('input', update);
      chartValueAxisInput.addEventListener('input', update);
      chartValueAxisPositionInput.addEventListener('input', update);
      chartDataLabelPositionInput.addEventListener('input', update);
      chartDataLabelSeparatorInput.addEventListener('input', update);
      chartDataLabelSeriesInput.addEventListener('input', update);
      chartDataLabelLegendInput.addEventListener('input', update);
      chartDataLabelLeaderInput.addEventListener('input', update);
      chartDataLabelPercentInput.addEventListener('input', update);
      chartDataLabelBubbleInput.addEventListener('input', update);
      chartSeriesInput.addEventListener('input', update);
      chartSeriesInvertInput.addEventListener('input', update);
      chartSeriesMarkerInput.addEventListener('input', update);
      chartSeriesExplosionInput.addEventListener('input', update);
      mediaTargetInput.addEventListener('input', update);
      commentInput.addEventListener('input', update);
      threadedCommentInput.addEventListener('input', update);
      threadedPersonInput.addEventListener('input', update);
      break;
    }
    case 'pptx': {
      const shapeText = officeDocument.slides[0]?.shapes[0]?.text ?? '';
      const notesText = officeDocument.slides[0]?.notesText ?? '';
      const commentText = officeDocument.slides[0]?.comments[0]?.text ?? '';
      const transitionType = officeDocument.slides[0]?.transition?.type ?? '';
      const transitionAdvanceOnClick = officeDocument.slides[0]?.transition?.advanceOnClick;
      const transitionAdvanceAfterMs = officeDocument.slides[0]?.transition?.advanceAfterMs ?? '';
      const firstTimingDuration = officeDocument.slides[0]?.timing?.nodes[0]?.duration ?? '';
      const firstTimingRepeatDuration = officeDocument.slides[0]?.timing?.nodes[0]?.repeatDuration ?? '';
      const firstTimingAutoReverse = officeDocument.slides[0]?.timing?.nodes[0]?.autoReverse;
      const firstTimingTrigger = officeDocument.slides[0]?.timing?.nodes[0]?.triggerEvent ?? '';
      const firstTimingEndTrigger = officeDocument.slides[0]?.timing?.nodes[0]?.endTriggerEvent ?? '';
      const firstTimingTarget = officeDocument.slides[0]?.timing?.nodes[0]?.targetShapeId ?? '';
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
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Advance on click</div>
          <input id="pptx-transition-click-input" value="${escapeHtml(transitionAdvanceOnClick === undefined ? '' : String(transitionAdvanceOnClick))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">Advance after ms</div>
          <input id="pptx-transition-after-input" value="${escapeHtml(String(transitionAdvanceAfterMs))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing duration</div>
          <input id="pptx-timing-duration-input" value="${escapeHtml(firstTimingDuration)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing repeat duration</div>
          <input id="pptx-timing-repeat-duration-input" value="${escapeHtml(firstTimingRepeatDuration)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing auto reverse</div>
          <input id="pptx-timing-auto-reverse-input" value="${escapeHtml(firstTimingAutoReverse === undefined ? '' : String(firstTimingAutoReverse))}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing trigger</div>
          <input id="pptx-timing-trigger-input" value="${escapeHtml(firstTimingTrigger)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing end trigger</div>
          <input id="pptx-timing-end-trigger-input" value="${escapeHtml(firstTimingEndTrigger)}" style="width: 100%; padding: 8px;" />
        </label>
        <label>
          <div style="font-size: 0.875rem; color: #475569; margin-bottom: 6px;">First timing target shape</div>
          <input id="pptx-timing-target-input" value="${escapeHtml(firstTimingTarget)}" style="width: 100%; padding: 8px;" />
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
      const transitionClickInput = window.document.getElementById('pptx-transition-click-input') as HTMLInputElement;
      const transitionAfterInput = window.document.getElementById('pptx-transition-after-input') as HTMLInputElement;
      const timingDurationInput = window.document.getElementById('pptx-timing-duration-input') as HTMLInputElement;
      const timingRepeatDurationInput = window.document.getElementById('pptx-timing-repeat-duration-input') as HTMLInputElement;
      const timingAutoReverseInput = window.document.getElementById('pptx-timing-auto-reverse-input') as HTMLInputElement;
      const timingTriggerInput = window.document.getElementById('pptx-timing-trigger-input') as HTMLInputElement;
      const timingEndTriggerInput = window.document.getElementById('pptx-timing-end-trigger-input') as HTMLInputElement;
      const timingTargetInput = window.document.getElementById('pptx-timing-target-input') as HTMLInputElement;
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
        setPresentationTransition(
          presentationEditor,
          0,
          transitionInput.value
            ? {
                type: transitionInput.value,
                advanceOnClick: transitionClickInput.value ? transitionClickInput.value === 'true' : undefined,
                advanceAfterMs: transitionAfterInput.value ? Number(transitionAfterInput.value) : undefined
              }
            : undefined
        );
        if (timingDurationInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, duration: timingDurationInput.value } : node)
          );
        }
        if (timingRepeatDurationInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, repeatDuration: timingRepeatDurationInput.value } : node)
          );
        }
        if (timingAutoReverseInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, autoReverse: timingAutoReverseInput.value === 'true' } : node)
          );
        }
        if (timingTriggerInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, triggerEvent: timingTriggerInput.value } : node)
          );
        }
        if (timingEndTriggerInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, endTriggerEvent: timingEndTriggerInput.value } : node)
          );
        }
        if (timingTargetInput.value && presentationEditor.document.slides[0]?.timing?.nodes.length) {
          const currentNodes = presentationEditor.document.slides[0].timing?.nodes ?? [];
          setPresentationTimingNodes(
            presentationEditor,
            0,
            currentNodes.map((node, index) => index === 0 ? { ...node, targetShapeId: timingTargetInput.value } : node)
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
      transitionClickInput.addEventListener('input', update);
      transitionAfterInput.addEventListener('input', update);
      timingDurationInput.addEventListener('input', update);
      timingRepeatDurationInput.addEventListener('input', update);
      timingAutoReverseInput.addEventListener('input', update);
      timingTriggerInput.addEventListener('input', update);
      timingEndTriggerInput.addEventListener('input', update);
      timingTargetInput.addEventListener('input', update);
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
