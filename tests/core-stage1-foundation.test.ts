import { describe, expect, it } from 'vitest';

import { clonePackageGraph, openPackage, summarizePackageProvenance, upsertRelationship } from '@ooxml/core';
import { serializeOfficeDocument } from '@ooxml/serializer';
import { parseDocx } from '@ooxml/docx';

import { createDocxFixture } from './fixture-builders';

describe('stage-1 preservation helpers', () => {
  it('serializes relationship mutations through .rels parts', async () => {
    const graph = clonePackageGraph(await openPackage(createDocxFixture()));

    upsertRelationship(graph, 'package', {
      id: 'rIdExt',
      type: 'https://example.com/relationships/external-link',
      target: 'https://example.com/resource',
      targetMode: 'External'
    });

    const reopened = await openPackage(serializeOfficeDocument(parseDocx(graph)));
    const relation = reopened.relationshipsBySource.package?.find((entry) => entry.id === 'rIdExt');

    expect(relation?.targetMode).toBe('External');
    expect(relation?.target).toBe('https://example.com/resource');
  });

  it('produces provenance summaries within stage-0 smoke thresholds for representative fixtures', async () => {
    const summary = summarizePackageProvenance(await openPackage(createDocxFixture()));

    expect(summary.xmlPartCount).toBeGreaterThan(0);
    expect(summary.estimatedRetainedBytes).toBeLessThanOrEqual(25 * 1024 * 1024);
    expect(summary.cloneTimeMs).toBeLessThanOrEqual(250);
  });
});
