import { openPackage, type PackageGraph } from '@ooxml/core';
import { inspectOfficeDocument, summarizePackageGraph, type OfficeDocumentSummary, type PackageGraphSummary } from '@ooxml/devtools';
import { createOfficeEditor, type EditableOfficeDocument, type OfficeEditor } from '@ooxml/editor';
import { parseDocx, type DocxDocument } from '@ooxml/docx';
import { parsePptx, type PresentationDocument } from '@ooxml/pptx';
import { mountOfficeDocument, renderOfficeDocumentToHtml, type RenderOptions } from '@ooxml/render';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseXlsx, type XlsxWorkbook } from '@ooxml/xlsx';

export type ParsedOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export interface BrowserSession {
  packageGraph: PackageGraph;
  document: ParsedOfficeDocument;
  packageSummary: PackageGraphSummary;
  documentSummary: OfficeDocumentSummary;
  renderToHtml: (options?: RenderOptions) => string;
  mount: (target: HTMLElement, options?: RenderOptions) => HTMLElement;
  createEditor(): OfficeEditor<EditableOfficeDocument>;
  save(): Blob;
}

export async function openOfficeDocument(input: ArrayBuffer | Uint8Array | Blob): Promise<ParsedOfficeDocument> {
  const graph = await openPackage(input);
  return parseOfficeDocument(graph);
}

export async function createBrowserSession(input: ArrayBuffer | Uint8Array | Blob): Promise<BrowserSession> {
  const packageGraph = await openPackage(input);
  const document = parseOfficeDocument(packageGraph);

  return {
    packageGraph,
    document,
    packageSummary: summarizePackageGraph(packageGraph),
    documentSummary: inspectOfficeDocument(document),
    renderToHtml(options = {}) {
      return renderOfficeDocumentToHtml(document, options);
    },
    mount(target, options = {}) {
      return mountOfficeDocument(document, target, options);
    },
    createEditor() {
      return createOfficeEditor(document);
    },
    save() {
      const bytes = Uint8Array.from(serializeOfficeDocument(document));
      return new Blob([bytes], {
        type: 'application/octet-stream'
      });
    }
  };
}

export function parseOfficeDocument(graph: PackageGraph): ParsedOfficeDocument {
  switch (graph.officeDocumentKind) {
    case 'docx':
      return parseDocx(graph);
    case 'xlsx':
      return parseXlsx(graph);
    case 'pptx':
      return parsePptx(graph);
    default:
      throw new Error(`Unsupported OOXML root document kind: ${graph.officeDocumentKind}`);
  }
}
