import { expect, test } from 'bun:test';
import { bodyPart } from './document';
import { formattingFingerprint } from './formatting-fingerprint';
import { canonicalOoxmlFingerprint } from '@docx-editor.dev/core/store';

test('formatting equivalence permits text-run splits but retains every native property and boundary', () => {
  const single = bodyPart(
    '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>'
  );
  const split = bodyPart(
    '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t>Alpha</w:t></w:r><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t xml:space="preserve"> beta</w:t></w:r></w:p>'
  );
  const original = canonicalOoxmlFingerprint(split);
  expect(formattingFingerprint(split)).toBe(formattingFingerprint(single));
  expect(canonicalOoxmlFingerprint(split)).toBe(original);
  for (const changed of [
    '<w:p><w:r><w:rPr><w:color w:val="112233"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:color w:val="112234" w:themeColor="accent1"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>',
    '<w:p><w:r w:rsidR="12345678"><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/><w:lang w:val="cs-CZ"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:p>',
    '<w:p><w:hyperlink w:anchor="target"><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t>Alpha beta</w:t></w:r></w:hyperlink></w:p>',
    '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t>Alpha</w:t><w:br/><w:t> beta</w:t></w:r></w:p>',
    '<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1"/></w:rPr><w:t>Alpha beta!</w:t></w:r></w:p>',
  ])
    expect(formattingFingerprint(bodyPart(changed))).not.toBe(formattingFingerprint(single));
});
