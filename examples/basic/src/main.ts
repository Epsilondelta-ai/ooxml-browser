import { createBrowserSession } from '@ooxml/browser';

const app = document.getElementById('app');

if (!app) {
  throw new Error('Missing #app mount point.');
}

app.innerHTML = `
  <section style="font-family: system-ui, sans-serif; max-width: 960px; margin: 0 auto; padding: 32px; display: grid; gap: 16px;">
    <header>
      <h1 style="margin: 0 0 8px;">OOXML Browser Example</h1>
      <p style="margin: 0; color: #475569;">Open a .docx, .xlsx, or .pptx file in the browser and preview the parsed render output.</p>
    </header>
    <label style="display: inline-flex; gap: 12px; align-items: center; background: #f8fafc; padding: 12px 16px; border-radius: 12px; border: 1px solid #cbd5e1;">
      <span>Select OOXML file</span>
      <input id="file-input" type="file" accept=".docx,.xlsx,.pptx" />
    </label>
    <p id="status" style="margin: 0; color: #334155;">Waiting for a file…</p>
    <pre id="summary" style="margin: 0; background: #0f172a; color: #e2e8f0; padding: 16px; border-radius: 12px; overflow: auto; min-height: 96px;"></pre>
    <div id="preview" style="background: white; border: 1px solid #cbd5e1; border-radius: 16px; padding: 20px; min-height: 200px;"></div>
  </section>
`;

const fileInput = document.getElementById('file-input') as HTMLInputElement;
const status = document.getElementById('status') as HTMLParagraphElement;
const summary = document.getElementById('summary') as HTMLPreElement;
const preview = document.getElementById('preview') as HTMLDivElement;

fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) {
    return;
  }

  status.textContent = `Opening ${file.name}…`;
  try {
    const session = await createBrowserSession(file);
    summary.textContent = JSON.stringify({
      packageSummary: session.packageSummary,
      documentSummary: session.documentSummary
    }, null, 2);
    session.mount(preview);
    status.textContent = `Loaded ${file.name}`;
  } catch (error) {
    preview.innerHTML = '';
    summary.textContent = '';
    status.textContent = error instanceof Error ? error.message : 'Failed to load file.';
  }
});
