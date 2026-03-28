import { clonePackageGraph, serializePackageGraph, updatePackagePartText } from '@ooxml/core';
import type { DocxComment, DocxDocument, DocxParagraph, DocxTable } from '@ooxml/docx';

export function serializeDocx(document: DocxDocument): Uint8Array {
  const graph = clonePackageGraph(document.packageGraph);
  const mainStory = document.stories.find((story) => story.kind === 'document');

  if (mainStory) {
    updatePackagePartText(
      graph,
      mainStory.uri,
      buildDocxStoryXml(mainStory.paragraphs, mainStory.tables),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
    );
  }

  const commentsUri = '/word/comments.xml';
  if (graph.parts[commentsUri]) {
    updatePackagePartText(
      graph,
      commentsUri,
      buildCommentsXml(document.comments),
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml'
    );
  }

  return serializePackageGraph(graph);
}

function buildDocxStoryXml(paragraphs: DocxParagraph[], tables: DocxTable[]): string {
  const paragraphXml = paragraphs.map((paragraph) => `<w:p>${paragraph.styleId ? `<w:pPr><w:pStyle w:val="${escapeXml(paragraph.styleId)}"/></w:pPr>` : ''}${paragraph.runs.map((run) => `<w:r>${run.bold || run.italic ? `<w:rPr>${run.bold ? '<w:b/>' : ''}${run.italic ? '<w:i/>' : ''}</w:rPr>` : ''}<w:t>${escapeXml(run.text)}</w:t></w:r>`).join('')}</w:p>`).join('');
  const tableXml = tables.map((table) => `<w:tbl>${table.rows.map((row) => `<w:tr>${row.cells.map((cell) => `<w:tc><w:p><w:r><w:t>${escapeXml(cell.text)}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`).join('');

  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphXml}${tableXml}</w:body></w:document>`;
}

function buildCommentsXml(comments: DocxComment[]): string {
  const commentXml = comments.map((comment) => `<w:comment w:id="${escapeXml(comment.id)}"${comment.author ? ` w:author="${escapeXml(comment.author)}"` : ''}><w:p><w:r><w:t>${escapeXml(comment.text)}</w:t></w:r></w:p></w:comment>`).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${commentXml}</w:comments>`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
