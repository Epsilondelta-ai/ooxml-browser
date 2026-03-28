import { XMLParser } from 'fast-xml-parser';

import { asArray } from './utils';
import type { ParsedXmlDocument } from './types';

type XmlNode = Record<string, unknown>;

const semanticParser = new XMLParser({
  attributeNamePrefix: '@_',
  commentPropName: '#comment',
  ignoreAttributes: false,
  parseAttributeValue: false,
  parseTagValue: false,
  preserveOrder: false,
  processEntities: false,
  trimValues: false
});

const tokenParser = new XMLParser({
  attributeNamePrefix: '@_',
  commentPropName: '#comment',
  ignoreAttributes: false,
  parseAttributeValue: false,
  parseTagValue: false,
  preserveOrder: true,
  processEntities: false,
  trimValues: false
});

export function parseXmlDocument(source: string): ParsedXmlDocument {
  return {
    source,
    tokens: tokenParser.parse(source) as unknown[],
    document: semanticParser.parse(source) as Record<string, unknown>
  };
}

export function xmlChildren<T = XmlNode>(node: unknown, key: string): T[] {
  if (!node || typeof node !== 'object') {
    return [];
  }

  const value = (node as XmlNode)[key] as T | T[] | undefined;
  return asArray(value);
}

export function xmlChild<T = XmlNode>(node: unknown, key: string): T | undefined {
  return xmlChildren<T>(node, key)[0];
}

export function xmlAttr(node: unknown, name: string): string | undefined {
  if (!node || typeof node !== 'object') {
    return undefined;
  }

  const raw = (node as XmlNode)[`@_${name}`];
  if (typeof raw === 'string') {
    return raw;
  }

  if (typeof raw === 'number' || typeof raw === 'boolean') {
    return String(raw);
  }

  return undefined;
}

export function xmlText(node: unknown): string {
  if (typeof node === 'string') {
    return node;
  }

  if (typeof node === 'number' || typeof node === 'boolean') {
    return String(node);
  }

  if (Array.isArray(node)) {
    return node.map((entry) => xmlText(entry)).join('');
  }

  if (!node || typeof node !== 'object') {
    return '';
  }

  const record = node as XmlNode;
  const segments: string[] = [];

  for (const [key, value] of Object.entries(record)) {
    if (key.startsWith('@_')) {
      continue;
    }

    if (key === '#text') {
      segments.push(xmlText(value));
      continue;
    }

    segments.push(xmlText(value));
  }

  return segments.join('');
}

export function findElementsByLocalName(node: unknown, localName: string): XmlNode[] {
  const matches: XmlNode[] = [];

  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const entry of value) {
        visit(entry);
      }
      return;
    }

    if (!value || typeof value !== 'object') {
      return;
    }

    for (const [key, child] of Object.entries(value as XmlNode)) {
      if (key.startsWith('@_') || key === '#text') {
        continue;
      }

      if (key.split(':').pop() === localName) {
        for (const match of asArray(child as XmlNode | XmlNode[])) {
          matches.push(match);
        }
      }

      visit(child);
    }
  };

  visit(node);
  return matches;
}
