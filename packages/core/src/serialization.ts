import { zipSync } from 'fflate';

import type { PackageGraph, PackagePart } from './types';
import { encodeText, normalizePartUri, partExtension } from './utils';

export function clonePackageGraph(graph: PackageGraph): PackageGraph {
  return {
    ...graph,
    contentTypes: {
      defaults: { ...graph.contentTypes.defaults },
      overrides: { ...graph.contentTypes.overrides }
    },
    diagnostics: [...graph.diagnostics],
    partOrder: [...graph.partOrder],
    relationshipsBySource: Object.fromEntries(
      Object.entries(graph.relationshipsBySource).map(([sourceUri, relationships]) => [
        sourceUri,
        relationships.map((relationship) => ({ ...relationship }))
      ])
    ),
    parts: Object.fromEntries(
      Object.entries(graph.parts).map(([uri, part]) => [uri, clonePackagePart(part)])
    )
  };
}

export function updatePackagePartText(graph: PackageGraph, uri: string, text: string, contentType?: string): void {
  const normalizedUri = normalizePartUri(uri);
  const existingPart = graph.parts[normalizedUri];
  const nextContentType = contentType ?? existingPart?.contentType ?? graph.contentTypes.defaults[partExtension(normalizedUri)] ?? 'application/xml';
  const data = encodeText(text);

  graph.parts[normalizedUri] = {
    uri: normalizedUri,
    extension: partExtension(normalizedUri),
    contentType: nextContentType,
    data,
    size: data.length,
    isXml: true,
    isRelationshipPart: normalizedUri.endsWith('.rels'),
    text,
    parsedXml: undefined
  };

  if (!graph.partOrder.includes(normalizedUri)) {
    graph.partOrder.push(normalizedUri);
    graph.partOrder.sort();
  }

  graph.contentTypes.overrides[normalizedUri] = nextContentType;
}

export function serializePackageGraph(graph: PackageGraph): Uint8Array {
  const archiveEntries: Record<string, Uint8Array> = {
    '[Content_Types].xml': encodeText(createContentTypesXml(graph))
  };

  for (const [uri, part] of Object.entries(graph.parts)) {
    archiveEntries[uri.slice(1)] = part.text !== undefined ? encodeText(part.text) : part.data;
  }

  return zipSync(archiveEntries);
}

export function createContentTypesXml(graph: PackageGraph): string {
  const defaults = Object.entries(graph.contentTypes.defaults)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([extension, contentType]) => `  <Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(contentType)}"/>`)
    .join('\n');
  const overrides = Object.entries(graph.contentTypes.overrides)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([partName, contentType]) => `  <Override PartName="${escapeXml(partName)}" ContentType="${escapeXml(contentType)}"/>`)
    .join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n${defaults}${defaults && overrides ? '\n' : ''}${overrides}\n</Types>`;
}

function clonePackagePart(part: PackagePart): PackagePart {
  return {
    ...part,
    data: new Uint8Array(part.data),
    parsedXml: part.parsedXml ? { ...part.parsedXml, tokens: [...part.parsedXml.tokens], document: structuredClone(part.parsedXml.document) } : undefined
  };
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
