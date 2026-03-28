import { getParsedXmlPart, relationshipById, relationshipsFor, findElementsByLocalName, xmlAttr, xmlChild, xmlChildren, xmlText, type PackageGraph } from '@ooxml/core';

import type { PresentationDocument, PresentationSlide } from './model';

export function parsePptx(graph: PackageGraph): PresentationDocument {
  const presentationUri = graph.rootDocumentUri ?? '/ppt/presentation.xml';
  const xml = getParsedXmlPart(graph, presentationUri);
  if (!xml) {
    throw new Error('Presentation part is missing.');
  }

  const presentation = xml.document['p:presentation'];
  const slideSize = xmlChild<Record<string, unknown>>(presentation, 'p:sldSz');
  const slidesRoot = xmlChild<Record<string, unknown>>(presentation, 'p:sldIdLst');

  const slides = xmlChildren<Record<string, unknown>>(slidesRoot, 'p:sldId').flatMap((slideReference) => {
    const relationshipId = xmlAttr(slideReference, 'r:id');
    const relationship = relationshipId ? relationshipById(graph, presentationUri, relationshipId) : undefined;
    if (!relationship?.resolvedTarget) {
      return [];
    }

    return [parseSlide(graph, relationship.resolvedTarget)];
  });

  return {
    kind: 'pptx',
    packageGraph: graph,
    slides,
    size: {
      cx: Number(xmlAttr(slideSize, 'cx') ?? '0'),
      cy: Number(xmlAttr(slideSize, 'cy') ?? '0')
    }
  };
}

function parseSlide(graph: PackageGraph, uri: string): PresentationSlide {
  const xml = getParsedXmlPart(graph, uri);
  if (!xml) {
    throw new Error(`Slide part ${uri} is missing.`);
  }

  const slide = xml.document['p:sld'];
  const commonSlideData = xmlChild<Record<string, unknown>>(slide, 'p:cSld');
  const shapeTree = xmlChild<Record<string, unknown>>(commonSlideData, 'p:spTree');
  const shapes = xmlChildren<Record<string, unknown>>(shapeTree, 'p:sp').map((shape) => {
    const nvSpPr = xmlChild<Record<string, unknown>>(shape, 'p:nvSpPr');
    const cNvPr = xmlChild<Record<string, unknown>>(nvSpPr, 'p:cNvPr');
    const textNodes = findElementsByLocalName(shape, 't');

    return {
      id: xmlAttr(cNvPr, 'id') ?? '',
      name: xmlAttr(cNvPr, 'name'),
      text: textNodes.map((node) => xmlText(node)).join('')
    };
  });

  const notesInfo = parseNotesInfo(graph, uri);
  const title = shapes.find((shape) => shape.text.trim())?.text ?? 'Slide';

  return {
    title,
    uri,
    notesUri: notesInfo.uri,
    notesText: notesInfo.text,
    shapes
  };
}

function parseNotesInfo(graph: PackageGraph, slideUri: string): { uri?: string; text: string } {
  const notesRelationship = relationshipsFor(graph, slideUri).find((relationship) => relationship.type.includes('/notesSlide'));
  if (!notesRelationship?.resolvedTarget) {
    return { text: '' };
  }

  const xml = getParsedXmlPart(graph, notesRelationship.resolvedTarget);
  if (!xml) {
    return { uri: notesRelationship.resolvedTarget, text: '' };
  }

  return {
    uri: notesRelationship.resolvedTarget,
    text: findElementsByLocalName(xml.document, 't').map((node) => xmlText(node)).join('')
  };
}
