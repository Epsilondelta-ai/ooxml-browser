import { readFile } from 'node:fs/promises';

import { describe, expect, it } from 'vitest';

describe('benchmark artifacts', () => {
  it('stores micro and representative benchmark results within current checkpoint thresholds', async () => {
    const report = JSON.parse(await readFile('benchmarks/reports/latest-benchmark-results.json', 'utf8')) as {
      suite: string;
      micro: Array<{ label: string; openMs: number; parseMs: number; renderMs: number; serializeMs: number }>;
      representative: Array<{ label: string; openMs: number; parseMs: number; renderMs: number; serializeMs: number }>;
    };

    expect(report.suite).toBe('ooxml-benchmarks');
    expect(report.micro.length).toBeGreaterThanOrEqual(3);
    expect(report.representative.length).toBeGreaterThanOrEqual(3);

    for (const result of [...report.micro, ...report.representative]) {
      expect(result.openMs).toBeLessThan(100);
      expect(result.parseMs).toBeLessThan(50);
      expect(result.renderMs).toBeLessThan(50);
      expect(result.serializeMs).toBeLessThan(50);
    }
  });
});
