import type { PackageGraph } from './types';
import { getParsedXmlPart } from './opc';

export interface PackageProvenancePartSummary {
  uri: string;
  textBytes: number;
  tokenCount: number;
}

export interface PackageProvenanceSummary {
  xmlPartCount: number;
  totalTextBytes: number;
  totalTokenCount: number;
  estimatedRetainedBytes: number;
  cloneTimeMs: number;
  parts: PackageProvenancePartSummary[];
}

export function summarizePackageProvenance(graph: PackageGraph): PackageProvenanceSummary {
  const parts = Object.values(graph.parts)
    .filter((part) => part.isXml)
    .map((part) => {
      const parsed = getParsedXmlPart(graph, part.uri);
      const tokenCount = Array.isArray(parsed?.tokens) ? parsed.tokens.length : 0;
      const textBytes = part.data.byteLength;

      return {
        uri: part.uri,
        textBytes,
        tokenCount
      } satisfies PackageProvenancePartSummary;
    });

  const payload = {
    xmlPartCount: parts.length,
    totalTextBytes: parts.reduce((sum, part) => sum + part.textBytes, 0),
    totalTokenCount: parts.reduce((sum, part) => sum + part.tokenCount, 0),
    parts
  };

  const now = globalThis.performance?.now.bind(globalThis.performance) ?? Date.now;
  const start = now();
  structuredClone(payload);
  const cloneTimeMs = now() - start;

  return {
    ...payload,
    estimatedRetainedBytes: payload.totalTextBytes + payload.totalTokenCount * 32,
    cloneTimeMs: Number(cloneTimeMs.toFixed(3))
  };
}
