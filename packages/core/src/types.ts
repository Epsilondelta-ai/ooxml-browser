export type OfficeDocumentKind = 'docx' | 'xlsx' | 'pptx' | 'unknown';
export type RelationshipTargetMode = 'Internal' | 'External';
export type RelationshipSourceUri = 'package' | string;

export interface Diagnostic {
  code: string;
  message: string;
  severity: 'info' | 'warning' | 'error';
  partUri?: string;
}

export interface ParsedXmlDocument {
  source: string;
  tokens: unknown[];
  document: Record<string, unknown>;
}

export interface ContentTypesMap {
  defaults: Record<string, string>;
  overrides: Record<string, string>;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode: RelationshipTargetMode;
  sourceUri: RelationshipSourceUri;
  resolvedTarget: string | null;
}

export interface PackagePart {
  uri: string;
  extension: string;
  contentType: string;
  data: Uint8Array;
  size: number;
  isXml: boolean;
  isRelationshipPart: boolean;
  text?: string;
  parsedXml?: ParsedXmlDocument;
}

export interface PackageGraph {
  parts: Record<string, PackagePart>;
  partOrder: string[];
  contentTypes: ContentTypesMap;
  relationshipsBySource: Record<string, Relationship[]>;
  rootDocumentUri: string | null;
  officeDocumentKind: OfficeDocumentKind;
  diagnostics: Diagnostic[];
}
