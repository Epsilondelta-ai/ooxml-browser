import { renderDocx } from './docx';
import { renderPresentation } from './pptx';
import { renderWorkbook } from './xlsx';

export type { RenderOptions, RenderableOfficeDocument } from './types';
import type { RenderOptions, RenderableOfficeDocument } from './types';

export function renderOfficeDocumentToHtml(document: RenderableOfficeDocument, options: RenderOptions = {}): string {
  switch (document.kind) {
    case 'docx':
      return renderDocx(document, options);
    case 'xlsx':
      return renderWorkbook(document, options);
    case 'pptx':
      return renderPresentation(document, options);
    default:
      return '<section class="ooxml-render ooxml-render--unknown">Unsupported document</section>';
  }
}

export function mountOfficeDocument(document: RenderableOfficeDocument, target: HTMLElement, options: RenderOptions = {}): HTMLElement {
  target.innerHTML = renderOfficeDocumentToHtml(document, options);
  return target;
}
