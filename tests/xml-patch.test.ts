import { describe, expect, it } from 'vitest';

import { replaceAttributeValue, replaceContainerTextByAttribute, replaceInnerTextByAttribute } from '@ooxml/core';

describe('shared xml patch helpers', () => {
  it('updates matching attributes while preserving unrelated ones', () => {
    const source = '<table id="1" name="Old" displayName="Old" ref="A1:B2" custom="keep"/>';
    const patched = replaceAttributeValue(source, { tagName: 'table', targetAttr: 'ref', newValue: 'A1:B3' });

    expect(patched).toContain('ref="A1:B3"');
    expect(patched).toContain('custom="keep"');
  });


  it('updates matching attributes by occurrence without touching siblings', () => {
    const source = '<comments><p:cm authorId="A"/><p:cm authorId="B"/></comments>';
    const patched = replaceAttributeValue(source, { tagName: 'p:cm', targetAttr: 'authorId', newValue: 'C', occurrence: 1 });

    expect(patched).toContain('authorId="A"');
    expect(patched).toContain('authorId="C"');
  });

  it('updates nested text inside a matching container without dropping siblings', () => {
    const source = '<comments><comment ref="B2"><meta keep="1"/><text><r><t>Old</t></r></text></comment></comments>';
    const patched = replaceInnerTextByAttribute(source, { containerTag: 'comment', keyAttr: 'ref', keyValue: 'B2', textTag: 't', newText: 'Updated' });

    expect(patched).toContain('<meta keep="1"/>');
    expect(patched).toContain('<t>Updated</t>');
  });

  it('updates direct container text while preserving outer attributes', () => {
    const source = '<definedNames><definedName name="SalesRange" localSheetId="0">Sheet1!$A$1:$B$2</definedName></definedNames>';
    const patched = replaceContainerTextByAttribute(source, { tagName: 'definedName', keyAttr: 'name', keyValue: 'SalesRange', newText: 'Sheet1!$A$1:$B$3' });

    expect(patched).toContain('localSheetId="0"');
    expect(patched).toContain('>Sheet1!$A$1:$B$3</definedName>');
  });
});
