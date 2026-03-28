import { replaceAttributeValue, replaceContainerTextByAttribute, replaceInnerTextByAttribute } from './xml-patch';

export type XmlPatchOperation =
  | {
      op: 'replaceAttribute';
      tagName: string;
      targetAttr: string;
      newValue: string;
      keyAttr?: string;
      keyValue?: string;
    }
  | {
      op: 'replaceText';
      containerTag: string;
      textTag: string;
      newText: string;
      keyAttr?: string;
      keyValue?: string;
      occurrence?: number;
    }
  | {
      op: 'replaceContainerText';
      tagName: string;
      newText: string;
      keyAttr?: string;
      keyValue?: string;
      occurrence?: number;
    };

export function applyXmlPatchPlan(source: string, operations: XmlPatchOperation[]): string {
  return operations.reduce((current, operation) => {
    if (operation.op === 'replaceAttribute') {
      return replaceAttributeValue(current, operation);
    }

    if (operation.op === 'replaceContainerText') {
      return replaceContainerTextByAttribute(current, operation);
    }

    return replaceInnerTextByAttribute(current, operation);
  }, source);
}
