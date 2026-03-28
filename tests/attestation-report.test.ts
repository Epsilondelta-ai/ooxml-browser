import { readFile } from 'node:fs/promises';

import { describe, expect, it } from 'vitest';

describe('attestation report artifacts', () => {
  it('tracks attestation status for every manifest-backed fixture', async () => {
    const report = JSON.parse(await readFile('benchmarks/reports/latest-attestation-report.json', 'utf8')) as {
      results: Array<{
        fixtureId: string;
        office: { status: string };
        libreOffice: { status: string };
      }>;
      summary: {
        office: Record<string, number>;
        libreOffice: Record<string, number>;
      };
    };

    expect(report.results.length).toBeGreaterThanOrEqual(14);
    for (const result of report.results) {
      expect(result.fixtureId.length).toBeGreaterThan(0);
      expect(result.office.status.length).toBeGreaterThan(0);
      expect(result.libreOffice.status.length).toBeGreaterThan(0);
    }
    expect(Object.values(report.summary.office).reduce((sum, count) => sum + count, 0)).toBe(report.results.length);
    expect(Object.values(report.summary.libreOffice).reduce((sum, count) => sum + count, 0)).toBe(report.results.length);
  });
});
