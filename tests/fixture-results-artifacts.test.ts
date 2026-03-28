import { readFile } from 'node:fs/promises';

import { describe, expect, it } from 'vitest';

describe('fixture result artifacts', () => {
  it('stores parser and edited round-trip evidence for every manifest-backed fixture', async () => {
    const report = JSON.parse(await readFile('benchmarks/reports/latest-fixture-results.json', 'utf8')) as {
      results: Array<{
        id: string;
        mutation: string;
        parserOpen: boolean;
        parserRoundTrip: boolean;
        editedRoundTrip: boolean;
        changedParts: string[];
        addedParts: string[];
        removedParts: string[];
        changedPartCount: number;
        preservedPartCount: number;
        totalOriginalPartCount: number;
      }>;
    };

    expect(report.results.length).toBeGreaterThanOrEqual(14);
    for (const result of report.results) {
      expect(result.mutation.length).toBeGreaterThan(0);
      expect(result.parserOpen).toBe(true);
      expect(result.parserRoundTrip).toBe(true);
      expect(result.editedRoundTrip).toBe(true);
      expect(Array.isArray(result.changedParts)).toBe(true);
      expect(Array.isArray(result.addedParts)).toBe(true);
      expect(Array.isArray(result.removedParts)).toBe(true);
      expect(result.changedPartCount).toBe(result.changedParts.length);
      expect(result.totalOriginalPartCount).toBeGreaterThan(0);
      expect(result.preservedPartCount).toBeGreaterThanOrEqual(0);
      expect(result.preservedPartCount + result.changedPartCount + result.removedParts.length).toBe(result.totalOriginalPartCount);
    }
  });

  it('prefers narrow SpreadsheetML mutations that avoid shared-string churn when possible', async () => {
    const report = JSON.parse(await readFile('benchmarks/reports/latest-fixture-results.json', 'utf8')) as {
      results: Array<{ id: string; changedParts: string[] }>;
    };

    const basic = report.results.find((result) => result.id === 'xlsx-basic');
    const styled = report.results.find((result) => result.id === 'xlsx-styled');

    expect(basic?.changedParts).toEqual(['/xl/worksheets/sheet1.xml']);
    expect(styled?.changedParts).toEqual(['/xl/worksheets/sheet1.xml']);
  });
});
