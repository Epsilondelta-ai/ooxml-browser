import { applyXmlPatchPlan, clonePackageGraph, serializePackageGraph, updatePackagePartText } from '@ooxml/core';
import { parseDocx, type DocxComment, type DocxDocument, type DocxParagraph, type DocxSection, type DocxStyle, type DocxTable } from '@ooxml/docx';

export function serializeDocx(document: DocxDocument): Uint8Array {
  const graph = clonePackageGraph(document.packageGraph);
  const originalDocument = parseDocx(document.packageGraph);

  for (const story of document.stories) {
    const contentType = story.kind === 'header'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
      : story.kind === 'footer'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
        : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

    updatePackagePartText(
      graph,
      story.uri,
      buildDocxStoryXml(story.paragraphs, story.tables, story.kind === 'document' ? document.sections[0] : undefined, story.kind, story.blocks, graph.parts[story.uri]?.text),
      contentType
    );
  }

  const commentsUri = '/word/comments.xml';
  if (graph.parts[commentsUri] && JSON.stringify(document.comments) !== JSON.stringify(originalDocument.comments)) {
    const existingSource = graph.parts[commentsUri].text;
    const nextSource = existingSource ? patchDocxCommentsXml(existingSource, document.comments) : buildCommentsXml(document.comments);
    updatePackagePartText(
      graph,
      commentsUri,
      nextSource,
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

function buildDocxStoryXml(paragraphs: DocxParagraph[], tables: DocxTable[], section: DocxSection | undefined, kind: 'document' | 'header' | 'footer', blocks: DocxDocument['stories'][number]['blocks'] = [], existingSource?: string): string {
  if (existingSource && canPatchDocxStory(kind, blocks)) {
    let next = existingSource;
    let paragraphOccurrence = 0;
    for (const block of blocks) {
      if (block.kind !== 'paragraph') {
        continue;
      }
      next = applyXmlPatchPlan(next, [{ op: 'replaceText', containerTag: 'w:p', occurrence: paragraphOccurrence, textTag: 'w:t', newText: block.paragraph.text }]);
      paragraphOccurrence += 1;
    }
    return next;
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


function patchDocxCommentsXml(source: string, comments: DocxComment[]): string {
  const operations = comments.flatMap((comment) => [
    { op: 'replaceAttribute' as const, tagName: 'w:comment', keyAttr: 'w:id', keyValue: comment.id, targetAttr: 'w:author', newValue: comment.author ?? '' },
    { op: 'replaceText' as const, containerTag: 'w:comment', keyAttr: 'w:id', keyValue: comment.id, textTag: 'w:t', newText: comment.text }
  ]);
  return applyXmlPatchPlan(source, operations);
}


function canPatchDocxStory(kind: 'document' | 'header' | 'footer', blocks: DocxDocument['stories'][number]['blocks']): boolean {
  return blocks.every((block) => block.kind === 'paragraph' || block.kind === 'table');
}
