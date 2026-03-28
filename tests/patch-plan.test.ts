import { describe, expect, it } from 'vitest';

import { applyXmlPatchPlan } from '@ooxml/core';

describe('shared xml patch plan', () => {
  it('applies attribute and text operations in one pass', () => {
    const source = '<root><table name="Old" ref="A1:B2"/><comment ref="B2"><t>Old</t></comment><definedName name="SalesRange">Sheet1!$A$1:$B$2</definedName></root>';
    const patched = applyXmlPatchPlan(source, [
      { op: 'replaceAttribute', tagName: 'table', targetAttr: 'ref', newValue: 'A1:B3' },
      { op: 'replaceText', containerTag: 'comment', keyAttr: 'ref', keyValue: 'B2', textTag: 't', newText: 'Updated' },
      { op: 'replaceContainerText', tagName: 'definedName', keyAttr: 'name', keyValue: 'SalesRange', newText: 'Sheet1!$A$1:$B$3' }
    ]);

    expect(patched).toContain('ref="A1:B3"');
    expect(patched).toContain('<t>Updated</t>');
    expect(patched).toContain('>Sheet1!$A$1:$B$3</definedName>');
  });
});
