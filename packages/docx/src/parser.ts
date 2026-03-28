import { getParsedXmlPart, relationshipsFor, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { DocxComment, DocxDocument, DocxParagraph, DocxRun, DocxStory, DocxStyle, DocxTable } from './model';

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
    comments: parseComments(graph, mainDocumentUri),
    styles: parseStyles(graph, mainDocumentUri)
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

function parseStyles(graph: PackageGraph, mainDocumentUri: string): Record<string, DocxStyle> {
  const stylesRelationship = relationshipsFor(graph, mainDocumentUri).find((relationship) => relationship.type.includes('/styles'));
  const stylesUri = stylesRelationship?.resolvedTarget ?? '/word/styles.xml';
  const xml = getParsedXmlPart(graph, stylesUri);
  if (!xml) {
    return {};
  }

  const root = xml.document['w:styles'];
  return Object.fromEntries(
    xmlChildren<Record<string, unknown>>(root, 'w:style').map((styleNode) => {
      const styleId = xmlAttr(styleNode, 'w:styleId') ?? xmlAttr(styleNode, 'styleId') ?? '';
      const typeValue = xmlAttr(styleNode, 'w:type') ?? xmlAttr(styleNode, 'type') ?? 'unknown';
      const nameNode = xmlChild<Record<string, unknown>>(styleNode, 'w:name');
      const basedOnNode = xmlChild<Record<string, unknown>>(styleNode, 'w:basedOn');
      const runProperties = xmlChild<Record<string, unknown>>(styleNode, 'w:rPr');
      const style: DocxStyle = {
        id: styleId,
        type: (typeValue === 'paragraph' || typeValue === 'character' || typeValue === 'table' || typeValue === 'numbering' ? typeValue : 'unknown'),
        name: xmlAttr(nameNode, 'w:val') ?? xmlAttr(nameNode, 'val'),
        basedOn: xmlAttr(basedOnNode, 'w:val') ?? xmlAttr(basedOnNode, 'val'),
        isDefault: (xmlAttr(styleNode, 'w:default') ?? xmlAttr(styleNode, 'default')) === '1',
        bold: xmlChild<Record<string, unknown>>(runProperties, 'w:b') !== undefined ? true : undefined,
        italic: xmlChild<Record<string, unknown>>(runProperties, 'w:i') !== undefined ? true : undefined
      };

      return [styleId, style] satisfies [string, DocxStyle];
    })
  );
}

export function resolveDocxStyle(document: DocxDocument, styleId?: string): DocxStyle | undefined {
  if (!styleId) {
    return undefined;
  }

  const style = document.styles[styleId];
  if (!style) {
    return undefined;
  }

  if (!style.basedOn || !document.styles[style.basedOn]) {
    return style;
  }

  const parent = resolveDocxStyle(document, style.basedOn);
  return {
    ...parent,
    ...style,
    bold: style.bold ?? parent?.bold,
    italic: style.italic ?? parent?.italic
  };
}
