import type { PackageGraph, Relationship, RelationshipSourceUri } from './types';
import { normalizePartUri, resolveTargetUri } from './utils';
import { updatePackagePartText } from './serialization';

const RELATIONSHIP_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';

export function relationshipPartUriForSource(sourceUri: RelationshipSourceUri): string {
  if (sourceUri === 'package') {
    return '/_rels/.rels';
  }

  const normalizedSource = normalizePartUri(sourceUri);
  const lastSlashIndex = normalizedSource.lastIndexOf('/');
  const directory = normalizedSource.slice(0, lastSlashIndex + 1);
  const fileName = normalizedSource.slice(lastSlashIndex + 1);
  return normalizePartUri(`${directory}_rels/${fileName}.rels`);
}

export function upsertRelationship(graph: PackageGraph, sourceUri: RelationshipSourceUri, relationship: Omit<Relationship, 'sourceUri' | 'resolvedTarget'>): Relationship[] {
  const existing = graph.relationshipsBySource[sourceUri] ?? [];
  const nextRelationship: Relationship = {
    ...relationship,
    sourceUri,
    resolvedTarget: resolveTargetUri(sourceUri, relationship.target, relationship.targetMode)
  };
  const merged = [
    ...existing.filter((entry) => entry.id !== relationship.id),
    nextRelationship
  ].sort((left, right) => left.id.localeCompare(right.id));

  setRelationshipsForSource(graph, sourceUri, merged);
  return merged;
}

export function setRelationshipsForSource(graph: PackageGraph, sourceUri: RelationshipSourceUri, relationships: Relationship[]): void {
  graph.relationshipsBySource[sourceUri] = relationships.map((relationship) => ({
    ...relationship,
    sourceUri,
    resolvedTarget: resolveTargetUri(sourceUri, relationship.target, relationship.targetMode)
  }));

  updatePackagePartText(
    graph,
    relationshipPartUriForSource(sourceUri),
    buildRelationshipsXml(graph.relationshipsBySource[sourceUri]),
    RELATIONSHIP_CONTENT_TYPE
  );
}

export function buildRelationshipsXml(relationships: Relationship[]): string {
  const body = relationships
    .map((relationship) => `<Relationship Id="${escapeXml(relationship.id)}" Type="${escapeXml(relationship.type)}" Target="${escapeXml(relationship.target)}"${relationship.targetMode === 'External' ? ' TargetMode="External"' : ''}/>`)
    .join('');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
