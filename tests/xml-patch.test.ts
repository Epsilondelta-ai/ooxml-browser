import { describe, expect, it } from 'vitest';

import { replaceAttributeValue, replaceInnerTextByAttribute } from '@ooxml/core';

describe('shared xml patch helpers', () => {
  it('updates matching attributes while preserving unrelated ones', () => {
    const source = '<table id="1" name="Old" displayName="Old" ref="A1:B2" custom="keep"/>';
    const patched = replaceAttributeValue(source, { tagName: 'table', targetAttr: 'ref', newValue: 'A1:B3' });

    expect(patched).toContain('ref="A1:B3"');
    expect(patched).toContain('custom="keep"');
  });

  it('updates nested text inside a matching container without dropping siblings', () => {
    const source = '<comments><comment ref="B2"><meta keep="1"/><text><r><t>Old</t></r></text></comment></comments>';
    const patched = replaceInnerTextByAttribute(source, { containerTag: 'comment', keyAttr: 'ref', keyValue: 'B2', textTag: 't', newText: 'Updated' });

    expect(patched).toContain('<meta keep="1"/>');
    expect(patched).toContain('<t>Updated</t>');
  });
});
