import { getParsedXmlPart, relationshipById, relationshipsFor, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { DocxAbstractNumbering, DocxBlock, DocxComment, DocxDocument, DocxHeaderFooterReference, DocxNumbering, DocxNumberingInstance, DocxNumberingLevel, DocxParagraph, DocxRevision, DocxRun, DocxSection, DocxStory, DocxStyle, DocxTable } from './model';

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
    styles: parseStyles(graph, mainDocumentUri),
    numbering: parseNumbering(graph, mainDocumentUri),
    sections: parseSections(graph, mainDocumentUri)
  };
}

function parseStory(graph: PackageGraph, uri: string, kind: DocxStory['kind']): DocxStory | undefined {
  const xml = getParsedXmlPart(graph, uri);
  if (!xml) {
    return undefined;
  }

  const root = xml.document['w:document'] ?? xml.document['w:hdr'] ?? xml.document['w:ftr'];
  const body = xmlChild<Record<string, unknown>>(root, 'w:body') ?? root;
  const paragraphs = xmlChildren<Record<string, unknown>>(body, 'w:p').map(parseParagraph);
  const tables = xmlChildren<Record<string, unknown>>(body, 'w:tbl').map(parseTable);
  const blocks = parseStoryBlocks(xml.tokens, kind, paragraphs, tables);

  return {
    kind,
    uri,
    blocks,
    paragraphs,
    tables
  };
}


function parseStoryBlocks(tokens: unknown[], kind: DocxStory['kind'], paragraphs: DocxParagraph[], tables: DocxTable[]): DocxBlock[] {
  const rootKey = kind === 'document' ? 'w:document' : kind === 'header' ? 'w:hdr' : 'w:ftr';
  const rootToken = tokens.find((token) => token && typeof token === 'object' && rootKey in (token as Record<string, unknown>)) as Record<string, unknown> | undefined;
  const rootChildren = rootToken?.[rootKey];
  const rootEntries = Array.isArray(rootChildren) ? rootChildren : [];
  const bodyEntries = kind === 'document'
    ? ((rootEntries.find((entry) => entry && typeof entry === 'object' && 'w:body' in (entry as Record<string, unknown>)) as Record<string, unknown> | undefined)?.['w:body'])
    : rootEntries;
  const entries = Array.isArray(bodyEntries) ? bodyEntries : [];

  const paragraphQueue = [...paragraphs];
  const tableQueue = [...tables];
  const blocks: DocxBlock[] = [];

  for (const entry of entries) {
    if (!entry || typeof entry !== 'object') {
      continue;
    }

    if ('w:p' in (entry as Record<string, unknown>)) {
      const paragraph = paragraphQueue.shift();
      if (paragraph) {
        blocks.push({ kind: 'paragraph', paragraph });
      }
    }

    if ('w:tbl' in (entry as Record<string, unknown>)) {
      const table = tableQueue.shift();
      if (table) {
        blocks.push({ kind: 'table', table });
      }
    }
  }

  if (blocks.length === 0) {
    return [
      ...paragraphs.map((paragraph) => ({ kind: 'paragraph', paragraph } as DocxBlock)),
      ...tables.map((table) => ({ kind: 'table', table } as DocxBlock))
    ];
  }

  return blocks;
}

function parseParagraph(node: Record<string, unknown>): DocxParagraph {
  const paragraphProperties = xmlChild<Record<string, unknown>>(node, 'w:pPr');
  const styleNode = xmlChild<Record<string, unknown>>(paragraphProperties, 'w:pStyle');
  const runs = xmlChildren<Record<string, unknown>>(node, 'w:r').map(parseRun);
  const revisions = [
    ...xmlChildren<Record<string, unknown>>(node, 'w:ins').map((revision) => parseRevision(revision, 'insertion')),
    ...xmlChildren<Record<string, unknown>>(node, 'w:del').map((revision) => parseRevision(revision, 'deletion'))
  ];

  const numPr = xmlChild<Record<string, unknown>>(paragraphProperties, 'w:numPr');
  const ilvlNode = xmlChild<Record<string, unknown>>(numPr, 'w:ilvl');
  const numIdNode = xmlChild<Record<string, unknown>>(numPr, 'w:numId');

  return {
    text: [...runs.map((run) => run.text), ...revisions.filter((revision) => revision.kind === 'insertion').map((revision) => revision.text)].join(''),
    styleId: xmlAttr(styleNode, 'w:val') ?? xmlAttr(styleNode, 'val'),
    numbering: numIdNode ? {
      numId: xmlAttr(numIdNode, 'w:val') ?? xmlAttr(numIdNode, 'val') ?? '',
      level: Number(xmlAttr(ilvlNode, 'w:val') ?? xmlAttr(ilvlNode, 'val') ?? '0')
    } : undefined,
    revisions,
    runs
  };
}


function parseRevision(node: Record<string, unknown>, kind: DocxRevision['kind']): DocxRevision {
  const runs = xmlChildren<Record<string, unknown>>(node, 'w:r').map((run) => {
    const textNodes = [
      ...xmlChildren<Record<string, unknown>>(run, 'w:t').map((entry) => xmlText(entry)),
      ...xmlChildren<Record<string, unknown>>(run, 'w:delText').map((entry) => xmlText(entry))
    ];
    return textNodes.join('');
  });

  return {
    kind,
    id: xmlAttr(node, 'w:id') ?? xmlAttr(node, 'id'),
    author: xmlAttr(node, 'w:author') ?? xmlAttr(node, 'author'),
    date: xmlAttr(node, 'w:date') ?? xmlAttr(node, 'date'),
    text: runs.join('')
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

function parseNumbering(graph: PackageGraph, mainDocumentUri: string): DocxNumbering {
  const numberingRelationship = relationshipsFor(graph, mainDocumentUri).find((relationship) => relationship.type.includes('/numbering'));
  const partUri = numberingRelationship?.resolvedTarget ?? (graph.parts['/word/numbering.xml'] ? '/word/numbering.xml' : undefined);
  if (!partUri) {
    return { abstractNums: {}, nums: {} };
  }

  const xml = getParsedXmlPart(graph, partUri);
  if (!xml) {
    return { partUri, abstractNums: {}, nums: {} };
  }

  const root = xml.document['w:numbering'];
  const abstractNums = Object.fromEntries(
    xmlChildren<Record<string, unknown>>(root, 'w:abstractNum').map((abstractNode) => {
      const id = xmlAttr(abstractNode, 'w:abstractNumId') ?? xmlAttr(abstractNode, 'abstractNumId') ?? '';
      const levels = Object.fromEntries(
        xmlChildren<Record<string, unknown>>(abstractNode, 'w:lvl').map((levelNode) => {
          const level = Number(xmlAttr(levelNode, 'w:ilvl') ?? xmlAttr(levelNode, 'ilvl') ?? '0');
          const numFmtNode = xmlChild<Record<string, unknown>>(levelNode, 'w:numFmt');
          const lvlTextNode = xmlChild<Record<string, unknown>>(levelNode, 'w:lvlText');
          const startNode = xmlChild<Record<string, unknown>>(levelNode, 'w:start');
          const item: DocxNumberingLevel = {
            level,
            format: xmlAttr(numFmtNode, 'w:val') ?? xmlAttr(numFmtNode, 'val') ?? 'decimal',
            text: xmlAttr(lvlTextNode, 'w:val') ?? xmlAttr(lvlTextNode, 'val') ?? '%1.',
            start: Number(xmlAttr(startNode, 'w:val') ?? xmlAttr(startNode, 'val') ?? '1')
          };
          return [level, item] satisfies [number, DocxNumberingLevel];
        })
      );
      const item: DocxAbstractNumbering = { id, levels };
      return [id, item] satisfies [string, DocxAbstractNumbering];
    })
  );

  const nums = Object.fromEntries(
    xmlChildren<Record<string, unknown>>(root, 'w:num').map((numNode) => {
      const id = xmlAttr(numNode, 'w:numId') ?? xmlAttr(numNode, 'numId') ?? '';
      const abstractNumIdNode = xmlChild<Record<string, unknown>>(numNode, 'w:abstractNumId');
      const item: DocxNumberingInstance = {
        id,
        abstractNumId: xmlAttr(abstractNumIdNode, 'w:val') ?? xmlAttr(abstractNumIdNode, 'val') ?? ''
      };
      return [id, item] satisfies [string, DocxNumberingInstance];
    })
  );

  return { partUri, abstractNums, nums };
}

export function resolveDocxNumbering(document: DocxDocument, paragraph: DocxParagraph): DocxNumberingLevel | undefined {
  if (!paragraph.numbering) {
    return undefined;
  }

  const num = document.numbering.nums[paragraph.numbering.numId];
  if (!num) {
    return undefined;
  }

  return document.numbering.abstractNums[num.abstractNumId]?.levels[paragraph.numbering.level];
}

function parseSections(graph: PackageGraph, mainDocumentUri: string): DocxSection[] {
  const xml = getParsedXmlPart(graph, mainDocumentUri);
  if (!xml) {
    return [];
  }

  const root = xml.document['w:document'];
  const body = xmlChild<Record<string, unknown>>(root, 'w:body');
  if (!body) {
    return [];
  }

  const sections: DocxSection[] = [];
  const bodySectPr = xmlChild<Record<string, unknown>>(body, 'w:sectPr');
  if (bodySectPr) {
    sections.push(parseSectionNode(graph, mainDocumentUri, bodySectPr, 'body'));
  }

  for (const paragraph of xmlChildren<Record<string, unknown>>(body, 'w:p')) {
    const paragraphProperties = xmlChild<Record<string, unknown>>(paragraph, 'w:pPr');
    const sectPr = xmlChild<Record<string, unknown>>(paragraphProperties, 'w:sectPr');
    if (sectPr) {
      sections.push(parseSectionNode(graph, mainDocumentUri, sectPr, 'paragraph'));
    }
  }

  return sections;
}

function parseSectionNode(graph: PackageGraph, mainDocumentUri: string, node: Record<string, unknown>, source: 'body' | 'paragraph'): DocxSection {
  const pageSizeNode = xmlChild<Record<string, unknown>>(node, 'w:pgSz');
  const pageMarginsNode = xmlChild<Record<string, unknown>>(node, 'w:pgMar');

  return {
    source,
    pageSize: pageSizeNode ? {
      width: Number(xmlAttr(pageSizeNode, 'w:w') ?? xmlAttr(pageSizeNode, 'w') ?? '0'),
      height: Number(xmlAttr(pageSizeNode, 'w:h') ?? xmlAttr(pageSizeNode, 'h') ?? '0')
    } : undefined,
    pageMargins: pageMarginsNode ? {
      top: Number(xmlAttr(pageMarginsNode, 'w:top') ?? xmlAttr(pageMarginsNode, 'top') ?? '0'),
      right: Number(xmlAttr(pageMarginsNode, 'w:right') ?? xmlAttr(pageMarginsNode, 'right') ?? '0'),
      bottom: Number(xmlAttr(pageMarginsNode, 'w:bottom') ?? xmlAttr(pageMarginsNode, 'bottom') ?? '0'),
      left: Number(xmlAttr(pageMarginsNode, 'w:left') ?? xmlAttr(pageMarginsNode, 'left') ?? '0')
    } : undefined,
    headerReferences: parseHeaderFooterReferences(graph, mainDocumentUri, node, 'w:headerReference'),
    footerReferences: parseHeaderFooterReferences(graph, mainDocumentUri, node, 'w:footerReference')
  };
}

function parseHeaderFooterReferences(graph: PackageGraph, mainDocumentUri: string, node: Record<string, unknown>, key: 'w:headerReference' | 'w:footerReference'): DocxHeaderFooterReference[] {
  return xmlChildren<Record<string, unknown>>(node, key).map((referenceNode) => {
    const relationshipId = xmlAttr(referenceNode, 'r:id') ?? '';
    const relationship = relationshipId ? relationshipById(graph, mainDocumentUri, relationshipId) : undefined;
    const rawType = xmlAttr(referenceNode, 'w:type') ?? xmlAttr(referenceNode, 'type') ?? 'default';
    return {
      type: rawType === 'first' || rawType === 'even' ? rawType : 'default',
      relationshipId,
      targetUri: relationship?.resolvedTarget ?? undefined
    };
  });
}
