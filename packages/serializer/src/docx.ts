import { applyXmlPatchPlan, clonePackageGraph, serializePackageGraph, updatePackagePartText, upsertRelationship } from '@ooxml/core';
import { parseDocx, type DocxComment, type DocxDocument, type DocxParagraph, type DocxSection, type DocxStyle, type DocxTable } from '@ooxml/docx';

export function serializeDocx(document: DocxDocument): Uint8Array {
  const graph = clonePackageGraph(document.packageGraph);
  const originalDocument = parseDocx(document.packageGraph);
  syncSectionReferenceRelationships(graph, originalDocument.sections[0], document.sections[0]);

  const originalStoriesByUri = new Map(originalDocument.stories.map((story) => [story.uri, story]));

  for (const story of document.stories) {
    syncStoryMediaRelationships(graph, originalStoriesByUri.get(story.uri), story);
    const contentType = story.kind === 'header'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
      : story.kind === 'footer'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

    updatePackagePartText(
      graph,
      story.uri,
      buildDocxStoryXml(story.paragraphs, story.tables, story.kind === 'document' ? document.sections[0] : undefined, story.kind, story.blocks, graph.parts[story.uri]?.text, originalStoriesByUri.get(story.uri), story.kind === 'document' ? originalDocument.sections[0] : undefined),
      contentType
    );
  }

  const commentsUri = '/word/comments.xml';
  if (document.comments.length > 0 && (!graph.parts[commentsUri] || JSON.stringify(document.comments) !== JSON.stringify(originalDocument.comments))) {
    ensureDocxCommentsRelationship(graph, document.packageGraph.rootDocumentUri ?? '/word/document.xml');
    const existingSource = graph.parts[commentsUri].text;
    const existingCommentCount = existingSource ? (existingSource.match(/<w:comment\b/g) ?? []).length : 0;
    const nextSource = existingSource && existingCommentCount === document.comments.length
      ? patchDocxCommentsXml(existingSource, document.comments)
      : buildCommentsXml(document.comments);
    updatePackagePartText(
      graph,
      commentsUri,
      nextSource,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    );
  }
  if (document.comments.length === 0 && graph.parts[commentsUri] && originalDocument.comments.length > 0) {
    updatePackagePartText(
      graph,
      commentsUri,
      buildCommentsXml([]),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    );
  }

  if (document.numbering.partUri && graph.parts[document.numbering.partUri] && JSON.stringify(document.numbering) !== JSON.stringify(originalDocument.numbering)) {
    updatePackagePartText(
      graph,
      document.numbering.partUri,
      buildNumberingXml(document.numbering),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
    );
  }

  const stylesUri = graph.parts['/word/styles.xml'] ? '/word/styles.xml' : undefined;
  if (stylesUri && JSON.stringify(document.styles) !== JSON.stringify(originalDocument.styles)) {
    updatePackagePartText(
      graph,
      stylesUri,
      buildStylesXml(document.styles),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
    );
  }

  return serializePackageGraph(graph);
}

function syncStoryMediaRelationships(graph: DocxDocument['packageGraph'], originalStory: DocxDocument['stories'][number] | undefined, story: DocxDocument['stories'][number]): void {
  for (const media of story.media) {
    const originalMedia = originalStory?.media.find((entry) => entry.relationshipId === media.relationshipId);
    if (!media.targetUri || media.targetUri === originalMedia?.targetUri) {
      continue;
    }

    upsertRelationship(graph, story.uri, {
      id: media.relationshipId,
      type: media.type === 'embeddedObject'
        ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'
        : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: relativeRelationshipTarget(story.uri, media.targetUri),
      targetMode: 'Internal'
    });
  }
}

function ensureDocxCommentsRelationship(graph: DocxDocument['packageGraph'], mainDocumentUri: string): void {
  const existing = graph.relationshipsBySource[mainDocumentUri]?.find((relationship) => relationship.type.includes('/comments'));
  upsertRelationship(graph, mainDocumentUri, {
    id: existing?.id ?? nextRelationshipId(graph.relationshipsBySource[mainDocumentUri] ?? []),
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    target: 'comments.xml',
    targetMode: 'Internal'
  });
}

function syncSectionReferenceRelationships(graph: DocxDocument['packageGraph'], originalSection: DocxSection | undefined, section: DocxSection | undefined): void {
  if (!section) {
    return;
  }

  syncReferenceGroup(graph, originalSection?.headerReferences ?? [], section.headerReferences, '/header');
  syncReferenceGroup(graph, originalSection?.footerReferences ?? [], section.footerReferences, '/footer');
}

function syncReferenceGroup(
  graph: DocxDocument['packageGraph'],
  originalReferences: DocxSection['headerReferences'],
  references: DocxSection['headerReferences'],
  typeFragment: '/header' | '/footer'
): void {
  for (const reference of references) {
    const originalReference = originalReferences.find((entry) => entry.relationshipId === reference.relationshipId);
    if (!reference.targetUri || reference.targetUri === originalReference?.targetUri) {
      continue;
    }

    upsertRelationship(graph, '/word/document.xml', {
      id: reference.relationshipId,
      type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships${typeFragment}`,
      target: relativeRelationshipTarget('/word/document.xml', reference.targetUri),
      targetMode: 'Internal'
    });
  }
}

function buildDocxStoryXml(paragraphs: DocxParagraph[], tables: DocxTable[], section: DocxSection | undefined, kind: 'document' | 'header' | 'footer', blocks: DocxDocument['stories'][number]['blocks'] = [], existingSource?: string, originalStory?: DocxDocument['stories'][number], originalSection?: DocxSection): string {
  if (existingSource && originalStory && JSON.stringify(originalStory.blocks) === JSON.stringify(blocks) && JSON.stringify(originalSection ?? null) === JSON.stringify(section ?? null)) {
    return existingSource;
  }

  if (existingSource && originalStory && canPatchDocxStory(kind, originalStory.blocks, blocks, originalSection, section)) {
    const operations = [];
    let paragraphOccurrence = 0;
    let originalParagraphIndex = 0;

    for (const block of blocks) {
      if (block.kind !== 'paragraph') {
        continue;
      }

      const originalBlock = originalStory.blocks.filter((candidate) => candidate.kind === 'paragraph')[originalParagraphIndex];
      if (originalBlock?.kind === 'paragraph' && originalBlock.paragraph.text !== block.paragraph.text) {
        operations.push({ op: 'replaceText' as const, containerTag: 'w:p', occurrence: paragraphOccurrence, textTag: 'w:t', newText: block.paragraph.text });
      }

      paragraphOccurrence += 1;
      originalParagraphIndex += 1;
    }

    return operations.length > 0 ? applyXmlPatchPlan(existingSource, operations) : existingSource;
  }

  const paragraphXmlFor = (paragraph: DocxParagraph) => {
    const paragraphProperties = [
      paragraph.styleId ? `<w:pStyle w:val="${escapeXml(paragraph.styleId)}"/>` : '',
      paragraph.numbering ? `<w:numPr><w:ilvl w:val="${paragraph.numbering.level}"/><w:numId w:val="${escapeXml(paragraph.numbering.numId)}"/></w:numPr>` : ''
    ].join('');

    const runsXml = paragraph.runs.map((run) => `<w:r>${run.bold || run.italic ? `<w:rPr>${run.bold ? '<w:b/>' : ''}${run.italic ? '<w:i/>' : ''}</w:rPr>` : ''}<w:t>${escapeXml(run.text)}</w:t></w:r>`).join('');
    const revisionsXml = paragraph.revisions.map((revision) => `<w:${revision.kind === 'insertion' ? 'ins' : 'del'}${revision.id ? ` w:id="${escapeXml(revision.id)}"` : ''}${revision.author ? ` w:author="${escapeXml(revision.author)}"` : ''}${revision.date ? ` w:date="${escapeXml(revision.date)}"` : ''}><w:r>${revision.kind === 'deletion' ? `<w:delText>${escapeXml(revision.text)}</w:delText>` : `<w:t>${escapeXml(revision.text)}</w:t>`}</w:r></w:${revision.kind === 'insertion' ? 'ins' : 'del'}>`).join('');

    return `<w:p>${paragraphProperties ? `<w:pPr>${paragraphProperties}</w:pPr>` : ''}${runsXml}${revisionsXml}</w:p>`;
  };
  const tableXmlFor = (table: DocxTable) => `<w:tbl>${table.rows.map((row) => `<w:tr>${row.cells.map((cell) => `<w:tc><w:p><w:r><w:t>${escapeXml(cell.text)}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`;
  const orderedBlocks = (blocks.length ? blocks : [
    ...paragraphs.map((paragraph) => ({ kind: 'paragraph', paragraph } as const)),
    ...tables.map((table) => ({ kind: 'table', table } as const))
  ]).map((block) => block.kind === 'paragraph' ? paragraphXmlFor(block.paragraph) : tableXmlFor(block.table)).join('');

  const sectionXml = kind === 'document' && section ? buildSectionXml(section) : '';
  const rootTag = kind === 'header' ? 'w:hdr' : kind === 'footer' ? 'w:ftr' : 'w:document';
  const bodyXml = kind === 'document' ? `<w:body>${orderedBlocks}${sectionXml}</w:body>` : `${orderedBlocks}`;
  return `<?xml version="1.0" encoding="UTF-8"?>\n<${rootTag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${bodyXml}</${rootTag}>`;
}

function buildSectionXml(section: DocxSection): string {
  const headerRefs = section.headerReferences.map((reference) => `<w:headerReference w:type="${escapeXml(reference.type)}" r:id="${escapeXml(reference.relationshipId)}"/>`).join('');
  const footerRefs = section.footerReferences.map((reference) => `<w:footerReference w:type="${escapeXml(reference.type)}" r:id="${escapeXml(reference.relationshipId)}"/>`).join('');
  const pageSize = section.pageSize ? `<w:pgSz w:w="${section.pageSize.width}" w:h="${section.pageSize.height}"/>` : '';
  const pageMargins = section.pageMargins ? `<w:pgMar w:top="${section.pageMargins.top}" w:right="${section.pageMargins.right}" w:bottom="${section.pageMargins.bottom}" w:left="${section.pageMargins.left}"/>` : '';
  return `<w:sectPr>${headerRefs}${footerRefs}${pageSize}${pageMargins}</w:sectPr>`;
}

function buildCommentsXml(comments: DocxComment[]): string {
  const commentXml = comments.map((comment) => `<w:comment w:id="${escapeXml(comment.id)}"${comment.author ? ` w:author="${escapeXml(comment.author)}"` : ''}><w:p><w:r><w:t>${escapeXml(comment.text)}</w:t></w:r></w:p></w:comment>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${commentXml}</w:comments>`;
}

function buildNumberingXml(numbering: DocxDocument['numbering']): string {
  const abstractNums = Object.values(numbering.abstractNums).map((abstractNum) => `<w:abstractNum w:abstractNumId="${escapeXml(abstractNum.id)}">${Object.values(abstractNum.levels).sort((a, b) => a.level - b.level).map((level) => `<w:lvl w:ilvl="${level.level}"><w:start w:val="${level.start}"/><w:numFmt w:val="${escapeXml(level.format)}"/><w:lvlText w:val="${escapeXml(level.text)}"/></w:lvl>`).join('')}</w:abstractNum>`).join('');
  const nums = Object.values(numbering.nums).map((num) => `<w:num w:numId="${escapeXml(num.id)}"><w:abstractNumId w:val="${escapeXml(num.abstractNumId)}"/></w:num>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abstractNums}${nums}</w:numbering>`;
}

function buildStylesXml(styles: Record<string, DocxStyle>): string {
  const styleXml = Object.values(styles).map((style) => `<w:style w:type="${escapeXml(style.type)}" w:styleId="${escapeXml(style.id)}"${style.isDefault ? ' w:default="1"' : ''}>${style.name ? `<w:name w:val="${escapeXml(style.name)}"/>` : ''}${style.basedOn ? `<w:basedOn w:val="${escapeXml(style.basedOn)}"/>` : ''}${style.bold || style.italic ? `<w:rPr>${style.bold ? '<w:b/>' : ''}${style.italic ? '<w:i/>' : ''}</w:rPr>` : ''}</w:style>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styleXml}</w:styles>`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}

function relativeRelationshipTarget(sourceUri: string, targetUri: string): string {
  const sourceSegments = sourceUri.replace(/^\//, '').split('/');
  sourceSegments.pop();
  const targetSegments = targetUri.replace(/^\//, '').split('/');

  while (sourceSegments.length > 0 && targetSegments.length > 0 && sourceSegments[0] === targetSegments[0]) {
    sourceSegments.shift();
    targetSegments.shift();
  }

  return `${sourceSegments.map(() => '..').join('/')}${sourceSegments.length ? '/' : ''}${targetSegments.join('/')}`;
}

function nextRelationshipId(relationships: Array<{ id: string }>): string {
  let candidateIndex = relationships.length + 1;
  let candidate = `rId${candidateIndex}`;
  const existingIds = new Set(relationships.map((relationship) => relationship.id));
  while (existingIds.has(candidate)) {
    candidateIndex += 1;
    candidate = `rId${candidateIndex}`;
  }
  return candidate;
}


function patchDocxCommentsXml(source: string, comments: DocxComment[]): string {
  const operations = comments.flatMap((comment) => [
    { op: 'replaceAttribute' as const, tagName: 'w:comment', keyAttr: 'w:id', keyValue: comment.id, targetAttr: 'w:author', newValue: comment.author ?? '' },
    { op: 'replaceText' as const, containerTag: 'w:comment', keyAttr: 'w:id', keyValue: comment.id, textTag: 'w:t', newText: comment.text }
  ]);
  return applyXmlPatchPlan(source, operations);
}


function canPatchDocxStory(kind: 'document' | 'header' | 'footer', originalBlocks: DocxDocument['stories'][number]['blocks'], blocks: DocxDocument['stories'][number]['blocks'], originalSection?: DocxSection, section?: DocxSection): boolean {
  if (kind !== 'document' && kind !== 'header' && kind !== 'footer') {
    return false;
  }

  if (originalBlocks.length !== blocks.length) {
    return false;
  }

  if (kind === 'document' && JSON.stringify(originalSection ?? null) !== JSON.stringify(section ?? null)) {
    return false;
  }

  return blocks.every((block, index) => {
    const originalBlock = originalBlocks[index];
    if (!originalBlock || block.kind !== originalBlock.kind) {
      return false;
    }

    if (block.kind !== 'paragraph' || originalBlock.kind !== 'paragraph') {
      return block.kind === 'table'
        && originalBlock.kind === 'table'
        && JSON.stringify(block.table) === JSON.stringify(originalBlock.table);
    }

    return block.paragraph.runs.length === 1
      && originalBlock.paragraph.runs.length === 1
      && block.paragraph.runs[0]?.bold === originalBlock.paragraph.runs[0]?.bold
      && block.paragraph.runs[0]?.italic === originalBlock.paragraph.runs[0]?.italic
      && block.paragraph.revisions.length === 0
      && originalBlock.paragraph.revisions.length === 0
      && block.paragraph.styleId === originalBlock.paragraph.styleId
      && JSON.stringify(block.paragraph.numbering ?? null) === JSON.stringify(originalBlock.paragraph.numbering ?? null);
  });
}
