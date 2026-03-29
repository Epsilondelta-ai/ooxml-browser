import { createBrowserSession } from '@ooxml/browser';

const app = document.getElementById('app');

if (!app) {
  throw new Error('Missing #app mount point.');
}

app.innerHTML = `
  <section id="example-root" style="font-family: system-ui, sans-serif; max-width: 1280px; margin: 0 auto; padding: 32px; display: grid; gap: 16px;">
    <style>
      #example-root.is-presentation-mode {
        max-width: 1400px;
      }

      #example-root.is-presentation-mode #example-intro p,
      #example-root.is-presentation-mode .preview-shell > h2 {
        display: none;
      }

      #example-root.is-presentation-mode #toolbar {
        padding: 8px 12px;
      }

      #example-root.is-presentation-mode #content-area {
        grid-template-columns: 1fr;
      }

      #example-root.is-presentation-mode #sidebar {
        order: 2;
        max-width: 960px;
        margin: 0 auto;
      }

      #example-root.is-presentation-mode .preview-shell {
        order: 1;
      }

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

      .slide-controls {
        display: flex;
        gap: 8px;
        align-items: center;
        flex-wrap: wrap;
      }

      .slide-controls button {
        border: 1px solid #cbd5e1;
        background: #fff;
        border-radius: 999px;
        padding: 6px 12px;
        font: inherit;
        cursor: pointer;
      }

      .slide-controls button:disabled {
        opacity: 0.5;
        cursor: not-allowed;
      }

      .slide-controls span {
        color: #475569;
        font-size: 0.95rem;
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
        max-width: 1280px;
        margin: 0 auto;
        border-radius: 18px;
        border: 1px solid #cbd5e1;
        background:
          linear-gradient(180deg, rgba(255,255,255,0.99), rgba(248,250,252,0.98)),
          radial-gradient(circle at top left, rgba(59,130,246,0.05), transparent 35%);
        box-shadow: 0 20px 40px rgba(15, 23, 42, 0.12);
        overflow: hidden;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-title-slide {
        background: linear-gradient(180deg, #1d2551 0%, #1a2148 100%);
        border-color: #1a2148;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-title-slide::before {
        content: '';
        position: absolute;
        left: 8%;
        top: 16%;
        width: 10%;
        height: 4px;
        border-radius: 999px;
        background: #3b82f6;
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
        justify-content: flex-start;
        padding: 0;
        border-radius: 0;
        border: none;
        background: transparent;
        box-shadow: none;
        overflow: hidden;
      }

      .preview-shell .ooxml-pptx-shape-content {
        white-space: pre-wrap;
        word-break: keep-all;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-content-slide .ooxml-pptx-shape.is-card {
        padding: 18px 18px 16px;
        border-radius: 12px;
        background: rgba(255,255,255,0.94);
        border: 1px solid #e5e7eb;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-content-slide .ooxml-pptx-shape.is-card::before {
        content: '';
        position: absolute;
        left: 0;
        top: 0;
        bottom: 0;
        width: 5px;
        background: var(--accent, #3b82f6);
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-content-slide .ooxml-pptx-shape.is-heading .ooxml-pptx-shape-content {
        font-size: 2rem;
        font-weight: 800;
        color: #1e293b;
        line-height: 1.15;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-content-slide .ooxml-pptx-shape.is-card .ooxml-pptx-shape-content {
        font-size: 1rem;
        color: #334155;
        line-height: 1.45;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-title-slide .ooxml-pptx-shape.kind-title .ooxml-pptx-shape-content {
        font-size: 4.8rem;
        font-weight: 800;
        color: #ffffff;
        line-height: 1.15;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-title-slide .ooxml-pptx-shape.kind-subtitle .ooxml-pptx-shape-content {
        font-size: 1.6rem;
        font-weight: 500;
        color: #93c5fd;
        line-height: 1.4;
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-title-slide .ooxml-pptx-shape.kind-footer .ooxml-pptx-shape-content {
        font-size: 0.95rem;
        color: rgba(255,255,255,0.7);
      }

      .preview-shell .ooxml-pptx-slide-canvas.is-content-slide .ooxml-pptx-shape.kind-body .ooxml-pptx-shape-content,
      .preview-shell .ooxml-pptx-slide-canvas.is-content-slide .ooxml-pptx-shape.kind-subtitle .ooxml-pptx-shape-content {
        font-size: 1.05rem;
        color: #334155;
        line-height: 1.45;
      }

      .preview-shell .ooxml-pptx-shape.is-media {
        padding: 14px 16px;
        border-radius: 14px;
        background: #0f172a;
        color: #e2e8f0;
        border: 1px solid #1e293b;
      }

      .preview-shell .ooxml-pptx-shape.is-media .ooxml-pptx-shape-content {
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
        font-size: 0.95rem;
      }

      .preview-shell .ooxml-pptx-shape.is-image-frame {
        padding: 0;
        border: none;
        background: transparent;
      }

      .preview-shell .ooxml-pptx-shape.is-image-frame img {
        width: 100%;
        height: 100%;
        object-fit: contain;
        display: block;
      }

      .preview-shell .ooxml-pptx-shape p,
      .preview-shell .ooxml-pptx-notes,
      .preview-shell .ooxml-pptx-comments,
      .preview-shell .ooxml-docx-comments {
        margin: 0;
      }
    </style>
    <header id="example-intro" style="display: grid; gap: 8px;">
      <h1 style="margin: 0;">OOXML file-input rendering example</h1>
      <p style="margin: 0; color: #475569;">Pick a <code>.docx</code>, <code>.xlsx</code>, or <code>.pptx</code> file with the file input below. The example opens it with <code>@ooxml/browser</code>, shows package/document summaries, visually previews the parsed output, and lets you download a round-tripped copy.</p>
    </header>
    <div id="toolbar" style="display: flex; gap: 12px; flex-wrap: wrap; align-items: center; background: #f8fafc; padding: 12px 16px; border-radius: 12px; border: 1px solid #cbd5e1;">
      <label style="display: inline-flex; gap: 12px; align-items: center;">
        <span>Select OOXML file</span>
        <input id="file-input" type="file" accept=".docx,.xlsx,.pptx" />
      </label>
      <button id="download-button" disabled>Download round-tripped file</button>
    </div>
    <p id="status" style="margin: 0; color: #334155;">Waiting for a file…</p>
    <div id="slide-controls" class="slide-controls" hidden>
      <button id="slide-prev-button" disabled>← Prev slide</button>
      <button id="slide-next-button" disabled>Next slide →</button>
      <span id="slide-indicator"></span>
    </div>
    <div id="content-area" style="display: grid; grid-template-columns: minmax(280px, 360px) 1fr; gap: 16px; align-items: start;">
      <div id="sidebar" style="display: grid; gap: 16px;">
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
const exampleRoot = document.getElementById('example-root') as HTMLElement;
const contentArea = document.getElementById('content-area') as HTMLDivElement;
const sidebar = document.getElementById('sidebar') as HTMLDivElement;
const downloadButton = document.getElementById('download-button') as HTMLButtonElement;
const status = document.getElementById('status') as HTMLParagraphElement;
const slideControls = document.getElementById('slide-controls') as HTMLDivElement;
const slidePrevButton = document.getElementById('slide-prev-button') as HTMLButtonElement;
const slideNextButton = document.getElementById('slide-next-button') as HTMLButtonElement;
const slideIndicator = document.getElementById('slide-indicator') as HTMLSpanElement;
const summary = document.getElementById('summary') as HTMLPreElement;
const htmlOutput = document.getElementById('html-output') as HTMLPreElement;
const preview = document.getElementById('preview') as HTMLDivElement;
let currentSession: Awaited<ReturnType<typeof createBrowserSession>> | null = null;
let currentFileName = 'document.ooxml';
let currentSlideIndex = 0;
const mediaUrlCache = new Map<string, string>();

fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) {
    return;
  }

  resetMediaUrlCache();
  currentFileName = file.name;
  currentSlideIndex = 0;
  status.textContent = `Opening ${file.name}…`;
  try {
    const session = await createBrowserSession(file);
    currentSession = session;
    renderLoadedDocument();
    downloadButton.disabled = false;
    status.textContent = `Loaded ${file.name}`;
  } catch (error) {
    currentSession = null;
    preview.innerHTML = '';
    summary.textContent = '';
    htmlOutput.textContent = '';
    slideControls.hidden = true;
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

slidePrevButton.addEventListener('click', () => {
  if (!currentSession || currentSession.document.kind !== 'pptx' || currentSlideIndex <= 0) {
    return;
  }

  currentSlideIndex -= 1;
  renderLoadedDocument();
});

slideNextButton.addEventListener('click', () => {
  if (!currentSession || currentSession.document.kind !== 'pptx') {
    return;
  }

  const maxIndex = currentSession.document.slides.length - 1;
  if (currentSlideIndex >= maxIndex) {
    return;
  }

  currentSlideIndex += 1;
  renderLoadedDocument();
});

function renderLoadedDocument(): void {
  if (!currentSession) {
    return;
  }

  const renderOptions = currentSession.document.kind === 'pptx'
    ? { activeSlideIndex: currentSlideIndex }
    : {};
  const html = currentSession.renderToHtml(renderOptions);
  summary.textContent = JSON.stringify({
    packageSummary: currentSession.packageSummary,
    documentSummary: currentSession.documentSummary,
    ...(currentSession.document.kind === 'pptx'
      ? {
          activeSlideIndex: currentSlideIndex,
          activeSlideTitle: currentSession.document.slides[currentSlideIndex]?.title ?? null
        }
      : {})
  }, null, 2);
  htmlOutput.textContent = html;
  currentSession.mount(preview, renderOptions);
  syncSlideControls();
  enhanceRenderedPreview();
}

function syncSlideControls(): void {
  if (!currentSession || currentSession.document.kind !== 'pptx') {
    exampleRoot.classList.remove('is-presentation-mode');
    contentArea.style.gridTemplateColumns = 'minmax(280px, 360px) 1fr';
    sidebar.style.display = 'grid';
    sidebar.style.order = '0';
    sidebar.style.maxWidth = '';
    sidebar.style.margin = '';
    slideControls.hidden = true;
    return;
  }

  exampleRoot.classList.add('is-presentation-mode');
  contentArea.style.gridTemplateColumns = '1fr';
  sidebar.style.display = 'none';
  const slideCount = currentSession.document.slides.length;
  slideControls.hidden = false;
  slidePrevButton.disabled = currentSlideIndex <= 0;
  slideNextButton.disabled = currentSlideIndex >= slideCount - 1;
  slideIndicator.textContent = `Slide ${currentSlideIndex + 1} / ${slideCount}: ${currentSession.document.slides[currentSlideIndex]?.title ?? ''}`;
}

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

  for (const selector of ['.ooxml-pptx-inheritance', '.ooxml-pptx-timing', '.ooxml-pptx-comments', '.ooxml-pptx-notes']) {
    const node = presentation.querySelector(selector) as HTMLElement | null;
    if (node) {
      node.style.display = 'none';
    }
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
  const backgroundColor = presentation.dataset.backgroundColor;
  const backgroundOpacity = Number(presentation.dataset.backgroundOpacity ?? 1);
  if (backgroundColor) {
    canvas.style.background = applyOpacity(backgroundColor, backgroundOpacity);
  }

  const accentPalette = ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444'];
  const titleLike = shapes.every((shape) => !shape.dataset.mediaType) && shapes.length <= 4;
  canvas.classList.toggle('is-title-slide', titleLike);
  canvas.classList.toggle('is-content-slide', !titleLike);
  const header = presentation.querySelector('header') as HTMLElement | null;
  if (header) {
    header.style.display = 'none';
    header.style.gap = '4px';
  }

  const orderedShapes = [...shapes].sort((left, right) => {
    const topDiff = Number(left.dataset.y ?? 0) - Number(right.dataset.y ?? 0);
    return topDiff !== 0 ? topDiff : Number(left.dataset.x ?? 0) - Number(right.dataset.x ?? 0);
  });

  for (const [index, shape] of orderedShapes.entries()) {
    if (shape.parentElement !== canvas) {
      canvas.appendChild(shape);
    }

    const name = shape.querySelector('h3')?.textContent?.trim() ?? '';
    const body = shape.querySelector('p')?.textContent?.trim() ?? '';
    const mediaType = shape.dataset.mediaType;
    const text = body || name || (mediaType === 'embeddedObject' ? '[embedded object]' : mediaType === 'image' ? '[image]' : '');
    shape.innerHTML = `<div class="ooxml-pptx-shape-content">${escapeHtmlText(text)}</div>`;

    const x = Number(shape.dataset.x ?? 0);
    const y = Number(shape.dataset.y ?? 0);
    const width = Number(shape.dataset.cx ?? 0);
    const height = Number(shape.dataset.cy ?? 0);
    const fillColor = shape.dataset.fillColor;
    const fillOpacity = Number(shape.dataset.fillOpacity ?? 1);
    const lineColor = shape.dataset.lineColor;
    const lineWidth = Number(shape.dataset.lineWidth ?? 0);
    const textColor = shape.dataset.textColor;
    const fontSizePt = Number(shape.dataset.fontSizePt ?? 0);
    const fontFamily = shape.dataset.fontFamily;
    const textAlign = shape.dataset.textAlign;

    shape.classList.remove('kind-title', 'kind-subtitle', 'kind-footer', 'kind-body', 'is-card', 'is-heading', 'is-media', 'is-image-frame');
    shape.style.removeProperty('--accent');
    shape.style.background = 'transparent';
    shape.style.border = 'none';

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

    if (fillColor) {
      shape.style.background = applyOpacity(fillColor, fillOpacity);
    }
    if (lineColor) {
      shape.style.border = `${Math.max(1, Math.round(lineWidth / 12700))}px solid ${lineColor}`;
    }

    if (mediaType) {
      const mediaUri = shape.dataset.mediaUri;
      const imageUrl = mediaType === 'image' && mediaUri ? getPackagePartObjectUrl(mediaUri) : null;
      if (imageUrl) {
        shape.classList.add('is-image-frame');
        shape.innerHTML = `<img src="${imageUrl}" alt="">`;
      } else {
        shape.classList.add('is-media');
        shape.style.setProperty('--accent', accentPalette[index % accentPalette.length]);
      }
      continue;
    }

    const content = shape.querySelector('.ooxml-pptx-shape-content') as HTMLElement | null;
    if (content) {
      content.style.color = textColor || '';
      content.style.fontFamily = fontFamily || '';
      content.style.textAlign = textAlign === 'ctr' ? 'center' : textAlign === 'r' ? 'right' : textAlign === 'just' ? 'justify' : 'left';
      if (fontSizePt > 0) {
        content.style.fontSize = `${Math.max(fontSizePt * 1.18, 14)}px`;
      }
      if (shape.dataset.fontBold === 'true') {
        content.style.fontWeight = '700';
      }
      if (shape.dataset.fontItalic === 'true') {
        content.style.fontStyle = 'italic';
      }
    }

    if (titleLike) {
      if (index === 0) {
        shape.classList.add('kind-title');
      } else if (index === orderedShapes.length - 1 && y / cy > 0.75) {
        shape.classList.add('kind-footer');
      } else {
        shape.classList.add('kind-subtitle');
      }
      continue;
    }

    const topRatio = y / cy;
    const widthRatio = width / cx;
    const heightRatio = height / cy;
    if (topRatio < 0.18 && widthRatio > 0.35) {
      shape.classList.add('kind-title', 'is-heading');
    } else if (widthRatio > 0.22 && heightRatio > 0.12) {
      shape.classList.add('kind-body', 'is-card');
      shape.style.setProperty('--accent', accentPalette[index % accentPalette.length]);
    } else {
      shape.classList.add('kind-body');
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

function escapeHtmlText(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function getPackagePartObjectUrl(uri: string): string | null {
  if (!currentSession) {
    return null;
  }

  const cached = mediaUrlCache.get(uri);
  if (cached) {
    return cached;
  }

  const part = currentSession.packageGraph.parts[uri];
  if (!part) {
    return null;
  }

  const blob = new Blob([part.data.slice()], { type: part.contentType || 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  mediaUrlCache.set(uri, url);
  return url;
}

function resetMediaUrlCache(): void {
  for (const url of mediaUrlCache.values()) {
    URL.revokeObjectURL(url);
  }
  mediaUrlCache.clear();
}

function applyOpacity(hexColor: string, opacity: number): string {
  if (!hexColor.startsWith('#') || hexColor.length !== 7) {
    return hexColor;
  }

  const red = Number.parseInt(hexColor.slice(1, 3), 16);
  const green = Number.parseInt(hexColor.slice(3, 5), 16);
  const blue = Number.parseInt(hexColor.slice(5, 7), 16);
  return `rgba(${red}, ${green}, ${blue}, ${Number.isFinite(opacity) ? opacity : 1})`;
}
