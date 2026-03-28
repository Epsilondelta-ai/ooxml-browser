import {
  getParsedXmlPart,
  relationshipsFor,
  type PackageGraph
} from '@ooxml/core';
import { xmlAttr, xmlChild, xmlChildren, xmlText } from '@ooxml/core';

export interface DocxRun {
  text: string;
  bold: boolean;
  italic: boolean;
}

export interface DocxParagraph {
  text: string;
  styleId?: string;
  runs: DocxRun[];
}

export interface DocxTableCell {
  text: string;
}

export interface DocxTableRow {
  cells: DocxTableCell[];
}

export interface DocxTable {
  rows: DocxTableRow[];
}

export interface DocxComment {
  id: string;
  author?: string;
  text: string;
}

export interface DocxStory {
  kind: 'document' | 'header' | 'footer';
  uri: string;
  paragraphs: DocxParagraph[];
  tables: DocxTable[];
}

export interface DocxDocument {
  kind: 'docx';
  packageGraph: PackageGraph;
  stories: DocxStory[];
  comments: DocxComment[];
}

export function parseDocx(graph: PackageGraph): DocxDocument {
  const mainDocumentUri = graph.rootDocumentUri ?? '/word/document.xml';
  const stories: DocxStory[] = [];

  const mainStory = parseStory(graph, mainDocumentUri, 'document');
  if (mainStory) {
    stories.push(mainStory);
  }

  for (const relationship of relationshipsFor(graph, mainDocumentUri)) {
    if (!relationship.resolvedTarget) {
      continue;
    }

    if (relationship.type.includes('/header')) {
      const story = parseStory(graph, relationship.resolvedTarget, 'header');
      if (story) {
        stories.push(story);
      }
    }

    if (relationship.type.includes('/footer')) {
      const story = parseStory(graph, relationship.resolvedTarget, 'footer');
      if (story) {
        stories.push(story);
      }
    }
  }

  return {
    kind: 'docx',
    packageGraph: graph,
    stories,
    comments: parseComments(graph, mainDocumentUri)
  };
}

function parseStory(graph: PackageGraph, uri: string, kind: DocxStory['kind']): DocxStory | undefined {
  const xml = getParsedXmlPart(graph, uri);
  if (!xml) {
    return undefined;
  }

  const root = xml.document['w:document'] ?? xml.document['w:hdr'] ?? xml.document['w:ftr'];
  const body = xmlChild<Record<string, unknown>>(root, 'w:body') ?? root;

  return {
    kind,
    uri,
    paragraphs: xmlChildren<Record<string, unknown>>(body, 'w:p').map(parseParagraph),
    tables: xmlChildren<Record<string, unknown>>(body, 'w:tbl').map(parseTable)
  };
}

function parseParagraph(node: Record<string, unknown>): DocxParagraph {
  const paragraphProperties = xmlChild<Record<string, unknown>>(node, 'w:pPr');
  const styleNode = xmlChild<Record<string, unknown>>(paragraphProperties, 'w:pStyle');
  const runs = xmlChildren<Record<string, unknown>>(node, 'w:r').map(parseRun);

  return {
    text: runs.map((run) => run.text).join(''),
    styleId: xmlAttr(styleNode, 'w:val') ?? xmlAttr(styleNode, 'val'),
    runs
  };
}

function parseRun(node: Record<string, unknown>): DocxRun {
  const runProperties = xmlChild<Record<string, unknown>>(node, 'w:rPr');
  const bold = xmlChild<Record<string, unknown>>(runProperties, 'w:b') !== undefined;
  const italic = xmlChild<Record<string, unknown>>(runProperties, 'w:i') !== undefined;
  const texts = xmlChildren<Record<string, unknown>>(node, 'w:t').map((entry) => xmlText(entry));
  const breaks = xmlChildren<Record<string, unknown>>(node, 'w:br').map(() => '\n');
  const tabs = xmlChildren<Record<string, unknown>>(node, 'w:tab').map(() => '\t');

  return {
    text: [...texts, ...breaks, ...tabs].join(''),
    bold,
    italic
  };
}

function parseTable(node: Record<string, unknown>): DocxTable {
  return {
    rows: xmlChildren<Record<string, unknown>>(node, 'w:tr').map((row) => ({
      cells: xmlChildren<Record<string, unknown>>(row, 'w:tc').map((cell) => ({
        text: xmlChildren<Record<string, unknown>>(cell, 'w:p').map((paragraph) => parseParagraph(paragraph).text).join('\n')
      }))
    }))
  };
}

function parseComments(graph: PackageGraph, mainDocumentUri: string): DocxComment[] {
  const commentRelationship = relationshipsFor(graph, mainDocumentUri).find((relationship) => relationship.type.includes('/comments'));
  if (!commentRelationship?.resolvedTarget) {
    return [];
  }

  const xml = getParsedXmlPart(graph, commentRelationship.resolvedTarget);
  if (!xml) {
    return [];
  }

  const root = xml.document['w:comments'];
  return xmlChildren<Record<string, unknown>>(root, 'w:comment').map((comment) => ({
    id: xmlAttr(comment, 'w:id') ?? xmlAttr(comment, 'id') ?? '',
    author: xmlAttr(comment, 'w:author') ?? xmlAttr(comment, 'author'),
    text: xmlChildren<Record<string, unknown>>(comment, 'w:p').map((paragraph) => parseParagraph(paragraph).text).join('\n')
  }));
}
