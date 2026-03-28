import { unzipSync } from 'fflate';

import type { ContentTypesMap, Diagnostic, OfficeDocumentKind, PackageGraph, PackagePart, Relationship, RelationshipSourceUri } from './types';
import { parseXmlDocument, xmlAttr, xmlChildren } from './xml';
import { decodeText, isXmlLikePath, normalizePartUri, partExtension, resolveTargetUri, toUint8Array } from './utils';

export interface OpenPackageOptions {
  maxEntryCount?: number;
  maxTotalUncompressedBytes?: number;
}

const DEFAULT_OPTIONS: Required<OpenPackageOptions> = {
  maxEntryCount: 10_000,
  maxTotalUncompressedBytes: 200 * 1024 * 1024
};

export async function openPackage(input: ArrayBuffer | Uint8Array | Blob, options: OpenPackageOptions = {}): Promise<PackageGraph> {
  const settings = { ...DEFAULT_OPTIONS, ...options };
  const bytes = await toUint8Array(input);
  const archive = unzipSync(bytes);
  const diagnostics: Diagnostic[] = [];
  const entryNames = Object.keys(archive);

  if (entryNames.length > settings.maxEntryCount) {
    throw new Error(`Package entry count ${entryNames.length} exceeds limit ${settings.maxEntryCount}`);
  }

  const totalUncompressedBytes = entryNames.reduce((sum, name) => sum + archive[name].length, 0);
  if (totalUncompressedBytes > settings.maxTotalUncompressedBytes) {
    throw new Error(`Package size ${totalUncompressedBytes} exceeds limit ${settings.maxTotalUncompressedBytes}`);
  }

  const contentTypesEntry = archive['[Content_Types].xml'];
  if (!contentTypesEntry) {
    throw new Error('Package is missing [Content_Types].xml');
  }

  const contentTypes = parseContentTypes(decodeText(contentTypesEntry));
  const parts: Record<string, PackagePart> = {};
  const partOrder = entryNames
    .map((name) => normalizePartUri(name))
    .sort();

  for (const name of entryNames) {
    const uri = normalizePartUri(name);
    if (uri === '/[Content_Types].xml') {
      continue;
    }

    const data = archive[name];
    const extension = partExtension(uri);
    const contentType = contentTypes.overrides[uri] ?? contentTypes.defaults[extension] ?? 'application/octet-stream';
    const isXml = isXmlLikePath(uri) || contentType.endsWith('+xml') || contentType.endsWith('/xml');

    parts[uri] = {
      uri,
      extension,
      contentType,
      data,
      size: data.length,
      isXml,
      isRelationshipPart: uri.endsWith('.rels'),
      text: isXml ? decodeText(data) : undefined
    };
  }

  const relationshipsBySource: Record<string, Relationship[]> = {};
  const packageRelationships = parseRelationshipsPart(parts['/_rels/.rels']?.text, 'package');
  relationshipsBySource.package = packageRelationships;

  for (const part of Object.values(parts)) {
    if (!part.isRelationshipPart || !part.text) {
      continue;
    }

    const sourceUri = relationshipSourceUriForPart(part.uri);
    relationshipsBySource[sourceUri] = parseRelationshipsPart(part.text, sourceUri);
  }

  const rootDocumentUri = packageRelationships.find((relationship) => relationship.type.includes('/officeDocument'))?.resolvedTarget ?? null;
  const officeDocumentKind = detectOfficeDocumentKind(rootDocumentUri, parts[rootDocumentUri ?? '']?.contentType);

  if (!rootDocumentUri) {
    diagnostics.push({
      code: 'opc.rootDocument.missing',
      message: 'Package relationships did not identify an officeDocument root.',
      severity: 'warning'
    });
  }

  return {
    parts,
    partOrder: partOrder.filter((uri) => uri !== '/[Content_Types].xml'),
    contentTypes,
    relationshipsBySource,
    rootDocumentUri,
    officeDocumentKind,
    diagnostics
  };
}

export function getPackagePart(graph: PackageGraph, uri: string): PackagePart | undefined {
  return graph.parts[normalizePartUri(uri)];
}

export function getPartText(graph: PackageGraph, uri: string): string | undefined {
  return getPackagePart(graph, uri)?.text;
}

export function getParsedXmlPart(graph: PackageGraph, uri: string): ReturnType<typeof parseXmlDocument> | undefined {
  const part = getPackagePart(graph, uri);
  if (!part || !part.text) {
    return undefined;
  }

  if (!part.parsedXml) {
    part.parsedXml = parseXmlDocument(part.text);
  }

  return part.parsedXml;
}

export function relationshipsFor(graph: PackageGraph, sourceUri: RelationshipSourceUri): Relationship[] {
  return graph.relationshipsBySource[sourceUri] ?? [];
}

export function relationshipById(graph: PackageGraph, sourceUri: RelationshipSourceUri, relationshipId: string): Relationship | undefined {
  return relationshipsFor(graph, sourceUri).find((relationship) => relationship.id === relationshipId);
}

export function detectOfficeDocumentKind(rootDocumentUri: string | null, contentType?: string): OfficeDocumentKind {
  const uri = rootDocumentUri ?? '';
  const type = contentType ?? '';

  if (uri.startsWith('/word/') || type.includes('wordprocessingml')) {
    return 'docx';
  }

  if (uri.startsWith('/xl/') || type.includes('spreadsheetml')) {
    return 'xlsx';
  }

  if (uri.startsWith('/ppt/') || type.includes('presentationml')) {
    return 'pptx';
  }

  return 'unknown';
}

function parseContentTypes(xmlSource: string): ContentTypesMap {
  const xml = parseXmlDocument(xmlSource);
  const types = xml.document.Types ?? xml.document['Types'];
  const defaults: Record<string, string> = {};
  const overrides: Record<string, string> = {};

  for (const entry of xmlChildren<Record<string, unknown>>(types, 'Default')) {
    const extension = xmlAttr(entry, 'Extension');
    const contentType = xmlAttr(entry, 'ContentType');
    if (extension && contentType) {
      defaults[extension.toLowerCase()] = contentType;
    }
  }

  for (const entry of xmlChildren<Record<string, unknown>>(types, 'Override')) {
    const partName = xmlAttr(entry, 'PartName');
    const contentType = xmlAttr(entry, 'ContentType');
    if (partName && contentType) {
      overrides[normalizePartUri(partName)] = contentType;
    }
  }

  return { defaults, overrides };
}

function parseRelationshipsPart(xmlSource: string | undefined, sourceUri: RelationshipSourceUri): Relationship[] {
  if (!xmlSource) {
    return [];
  }

  const xml = parseXmlDocument(xmlSource);
  const relationshipsRoot = xml.document.Relationships ?? xml.document['Relationships'];

  return xmlChildren<Record<string, unknown>>(relationshipsRoot, 'Relationship').map((entry) => {
    const id = xmlAttr(entry, 'Id') ?? '';
    const type = xmlAttr(entry, 'Type') ?? '';
    const target = xmlAttr(entry, 'Target') ?? '';
    const targetMode = (xmlAttr(entry, 'TargetMode') === 'External' ? 'External' : 'Internal') as Relationship['targetMode'];

    return {
      id,
      type,
      target,
      targetMode,
      sourceUri,
      resolvedTarget: resolveTargetUri(sourceUri, target, targetMode)
    } satisfies Relationship;
  });
}

function relationshipSourceUriForPart(relationshipPartUri: string): string {
  const normalized = normalizePartUri(relationshipPartUri);
  if (normalized === '/_rels/.rels') {
    return 'package';
  }

  const sourceUri = normalized
    .replace('/_rels/', '/')
    .replace(/\.rels$/, '');

  return normalizePartUri(sourceUri);
}
