import { createBrowserSession } from '@ooxml/browser';

const app = document.getElementById('app');

if (!app) {
  throw new Error('Missing #app mount point.');
}

app.innerHTML = `
  <section style="font-family: system-ui, sans-serif; max-width: 1100px; margin: 0 auto; padding: 32px; display: grid; gap: 16px;">
    <style>
      .preview-shell .ooxml-render {
        color: #0f172a;
        line-height: 1.5;
      }

      .preview-shell .ooxml-render--docx {
        max-width: 720px;
        margin: 0 auto;
        padding: 40px 48px;
        background: #fff;
        box-shadow: 0 12px 32px rgba(15, 23, 42, 0.08);
        border-radius: 12px;
      }

      .preview-shell .ooxml-docx-paragraph {
        margin: 0 0 1rem;
      }

      .preview-shell .ooxml-docx-table,
      .preview-shell .ooxml-xlsx-grid {
        width: 100%;
        border-collapse: collapse;
        background: #fff;
      }

      .preview-shell .ooxml-docx-table td,
      .preview-shell .ooxml-xlsx-grid td,
      .preview-shell .ooxml-xlsx-grid th {
        border: 1px solid #cbd5e1;
        padding: 8px 10px;
        vertical-align: top;
      }

      .preview-shell .ooxml-xlsx-grid th {
        background: #f8fafc;
        font-weight: 600;
        text-align: left;
      }

      .preview-shell .ooxml-render--xlsx {
        display: grid;
        gap: 12px;
      }

      .preview-shell .ooxml-xlsx-grid-shell {
        overflow: auto;
        border: 1px solid #cbd5e1;
        border-radius: 12px;
        background: #fff;
      }

      .preview-shell .ooxml-xlsx-charts,
      .preview-shell .ooxml-xlsx-media,
      .preview-shell .ooxml-xlsx-tables,
      .preview-shell .ooxml-xlsx-frozen-pane,
      .preview-shell .ooxml-xlsx-selection,
      .preview-shell .ooxml-xlsx-page-margins,
      .preview-shell .ooxml-xlsx-page-setup,
      .preview-shell .ooxml-xlsx-merged-ranges {
        margin: 0;
        padding: 10px 12px;
        border-radius: 10px;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        font-size: 0.95rem;
      }

      .preview-shell .ooxml-render--pptx {
        display: grid;
        gap: 16px;
      }

      .preview-shell .ooxml-xlsx-grid thead th {
        position: sticky;
        top: 0;
        z-index: 1;
      }

      .preview-shell .ooxml-pptx-slide-canvas {
        position: relative;
        width: 100%;
        max-width: 960px;
        margin: 0 auto;
        border-radius: 18px;
        border: 1px solid #cbd5e1;
        background:
          linear-gradient(180deg, rgba(255,255,255,0.98), rgba(248,250,252,0.98)),
          radial-gradient(circle at top left, rgba(59,130,246,0.08), transparent 35%);
        box-shadow: 0 20px 40px rgba(15, 23, 42, 0.12);
        overflow: hidden;
      }

      .preview-shell .ooxml-render--pptx > header,
      .preview-shell .ooxml-pptx-inheritance,
      .preview-shell .ooxml-pptx-timing,
      .preview-shell .ooxml-pptx-comments,
      .preview-shell .ooxml-pptx-notes {
        margin: 0;
        padding: 14px 16px;
        border-radius: 12px;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
      }

      .preview-shell .ooxml-pptx-shape {
        position: absolute;
        display: flex;
        flex-direction: column;
        justify-content: center;
        padding: 16px;
        border-radius: 14px;
        border: 1px solid #cbd5e1;
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
        overflow: hidden;
      }

      .preview-shell .ooxml-pptx-shape h3 {
        margin: 0 0 8px;
      }

      .preview-shell .ooxml-pptx-shape p,
      .preview-shell .ooxml-pptx-notes,
      .preview-shell .ooxml-pptx-comments,
      .preview-shell .ooxml-docx-comments {
        margin: 0;
      }
    </style>
    <header style="display: grid; gap: 8px;">
      <h1 style="margin: 0;">OOXML file-input rendering example</h1>
      <p style="margin: 0; color: #475569;">Pick a <code>.docx</code>, <code>.xlsx</code>, or <code>.pptx</code> file with the file input below. The example opens it with <code>@ooxml/browser</code>, shows package/document summaries, visually previews the parsed output, and lets you download a round-tripped copy.</p>
    </header>
    <div style="display: flex; gap: 12px; flex-wrap: wrap; align-items: center; background: #f8fafc; padding: 12px 16px; border-radius: 12px; border: 1px solid #cbd5e1;">
      <label style="display: inline-flex; gap: 12px; align-items: center;">
        <span>Select OOXML file</span>
        <input id="file-input" type="file" accept=".docx,.xlsx,.pptx" />
      </label>
      <button id="download-button" disabled>Download round-tripped file</button>
    </div>
    <p id="status" style="margin: 0; color: #334155;">Waiting for a file…</p>
    <div style="display: grid; grid-template-columns: minmax(280px, 360px) 1fr; gap: 16px; align-items: start;">
      <div style="display: grid; gap: 16px;">
        <section style="background: #0f172a; color: #e2e8f0; padding: 16px; border-radius: 16px;">
          <h2 style="margin-top: 0;">Session summary</h2>
          <pre id="summary" style="margin: 0; overflow: auto; min-height: 120px; white-space: pre-wrap;"></pre>
        </section>
        <details style="background: #f8fafc; border: 1px solid #cbd5e1; padding: 16px; border-radius: 16px;">
          <summary style="cursor: pointer; font-weight: 600;">Debug: generated HTML source</summary>
          <pre id="html-output" style="margin: 12px 0 0; overflow: auto; min-height: 160px; white-space: pre-wrap;"></pre>
        </details>
      </div>
      <section class="preview-shell" style="background: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 16px; padding: 20px; min-height: 280px;">
        <h2 style="margin-top: 0;">Visual preview</h2>
        <div id="preview"></div>
      </section>
    </div>
  </section>
`;

const fileInput = document.getElementById('file-input') as HTMLInputElement;
const downloadButton = document.getElementById('download-button') as HTMLButtonElement;
const status = document.getElementById('status') as HTMLParagraphElement;
const summary = document.getElementById('summary') as HTMLPreElement;
const htmlOutput = document.getElementById('html-output') as HTMLPreElement;
const preview = document.getElementById('preview') as HTMLDivElement;
let currentSession: Awaited<ReturnType<typeof createBrowserSession>> | null = null;
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
    const html = session.renderToHtml();
    summary.textContent = JSON.stringify({
      packageSummary: session.packageSummary,
      documentSummary: session.documentSummary
    }, null, 2);
    htmlOutput.textContent = html;
    session.mount(preview);
    enhanceRenderedPreview();
    downloadButton.disabled = false;
    status.textContent = `Loaded ${file.name}`;
  } catch (error) {
    currentSession = null;
    preview.innerHTML = '';
    summary.textContent = '';
    htmlOutput.textContent = '';
    downloadButton.disabled = true;
    status.textContent = error instanceof Error ? error.message : 'Failed to load file.';
  }
});

downloadButton.addEventListener('click', () => {
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
  status.textContent = `Downloaded round-tripped ${currentFileName}`;
});

function enhanceRenderedPreview(): void {
  enhanceWorkbookPreview();
  enhancePresentationPreview();
}

function enhanceWorkbookPreview(): void {
  const workbook = preview.querySelector('.ooxml-render--xlsx');
  const grid = workbook?.querySelector('.ooxml-xlsx-grid') as HTMLTableElement | null;
  if (!workbook || !grid) {
    return;
  }

  if (!grid.tHead) {
    const bodyRows = Array.from(grid.tBodies[0]?.rows ?? []);
    const maxColumns = bodyRows.reduce((count, row) => Math.max(count, row.cells.length - 1), 0);
    if (maxColumns > 0) {
      const head = grid.createTHead();
      const row = head.insertRow();
      row.insertCell().outerHTML = '<th scope="col"></th>';
      for (let index = 0; index < maxColumns; index += 1) {
        row.insertCell().outerHTML = `<th scope="col">${columnLabel(index)}</th>`;
      }
    }
  }

  if (!grid.parentElement?.classList.contains('ooxml-xlsx-grid-shell')) {
    const shell = document.createElement('div');
    shell.className = 'ooxml-xlsx-grid-shell';
    grid.parentElement?.insertBefore(shell, grid);
    shell.appendChild(grid);
  }
}

function enhancePresentationPreview(): void {
  const presentation = preview.querySelector('.ooxml-render--pptx') as HTMLElement | null;
  if (!presentation) {
    return;
  }

  const shapes = Array.from(presentation.querySelectorAll('.ooxml-pptx-shape')) as HTMLElement[];
  if (!shapes.length) {
    return;
  }

  let canvas = presentation.querySelector('.ooxml-pptx-slide-canvas') as HTMLDivElement | null;
  if (!canvas) {
    canvas = document.createElement('div');
    canvas.className = 'ooxml-pptx-slide-canvas';
    presentation.querySelector('header')?.insertAdjacentElement('afterend', canvas);
  }

  const cx = Number(presentation.dataset.presentationCx ?? 0) || 9144000;
  const cy = Number(presentation.dataset.presentationCy ?? 0) || 6858000;
  canvas.style.aspectRatio = `${cx} / ${cy}`;
  canvas.style.minHeight = '320px';

  for (const shape of shapes) {
    if (shape.parentElement !== canvas) {
      canvas.appendChild(shape);
    }

    const x = Number(shape.dataset.x ?? 0);
    const y = Number(shape.dataset.y ?? 0);
    const width = Number(shape.dataset.cx ?? 0);
    const height = Number(shape.dataset.cy ?? 0);

    if (width > 0 && height > 0) {
      shape.style.left = `${(x / cx) * 100}%`;
      shape.style.top = `${(y / cy) * 100}%`;
      shape.style.width = `${(width / cx) * 100}%`;
      shape.style.height = `${(height / cy) * 100}%`;
    } else {
      shape.style.left = '5%';
      shape.style.top = '5%';
      shape.style.width = '40%';
      shape.style.height = '20%';
    }
  }
}

function columnLabel(index: number): string {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}
