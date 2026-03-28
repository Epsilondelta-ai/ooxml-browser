import type { PackageGraph, Relationship } from '@ooxml/core';
import type { DocxDocument } from '@ooxml/docx';
import type { PresentationDocument } from '@ooxml/pptx';
import type { XlsxWorkbook } from '@ooxml/xlsx';

export type InspectableOfficeDocument = DocxDocument | XlsxWorkbook | PresentationDocument;

export interface PackageGraphSummary {
  officeDocumentKind: PackageGraph['officeDocumentKind'];
  rootDocumentUri: string | null;
  partCount: number;
  xmlPartCount: number;
  relationshipCount: number;
  contentTypes: string[];
}

export interface OfficeDocumentSummary {
  kind: InspectableOfficeDocument['kind'];
  primaryUnits: number;
  details: Record<string, number | string>;
}

export function summarizePackageGraph(graph: PackageGraph): PackageGraphSummary {
  const parts = Object.values(graph.parts);
  const relationshipCount = Object.values(graph.relationshipsBySource).reduce((sum, relationships) => sum + relationships.length, 0);

  return {
    officeDocumentKind: graph.officeDocumentKind,
    rootDocumentUri: graph.rootDocumentUri,
    partCount: parts.length,
    xmlPartCount: parts.filter((part) => part.isXml).length,
    relationshipCount,
    contentTypes: Array.from(new Set(parts.map((part) => part.contentType))).sort()
  };
}

export function inspectOfficeDocument(document: InspectableOfficeDocument): OfficeDocumentSummary {
  switch (document.kind) {
    case 'docx':
      return {
        kind: document.kind,
        primaryUnits: document.stories.length,
        details: {
          stories: document.stories.length,
          comments: document.comments.length,
          paragraphs: document.stories.reduce((sum, story) => sum + story.paragraphs.length, 0)
        }
      };
    case 'xlsx':
      return {
        kind: document.kind,
        primaryUnits: document.sheets.length,
        details: {
          sheets: document.sheets.length,
          sharedStrings: document.sharedStrings.length,
          cells: document.sheets.reduce((sum, sheet) => sum + sheet.rows.reduce((rowSum, row) => rowSum + row.cells.length, 0), 0)
        }
      };
    case 'pptx':
      return {
        kind: document.kind,
        primaryUnits: document.slides.length,
        details: {
          slides: document.slides.length,
          shapes: document.slides.reduce((sum, slide) => sum + slide.shapes.length, 0)
        }
      };
  }
}

export function flattenRelationships(graph: PackageGraph): Relationship[] {
  return Object.values(graph.relationshipsBySource).flat();
}
