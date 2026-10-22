import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { blockIdFromTag } from '../../src/blocks/payload-opc-codec';
import { GENERATED_LABEL_TAG } from '../../src/projection/block-tags';
import { W } from './document';
import { propertyEditor } from './property-execution';
import { groupFixture } from './review-groups';
import { proposeMixedBulk } from './bulk-decisions';

export interface ExternalLabelReview {
  readonly decision: 'accept' | 'reject';
  readonly only?: 'ins' | 'del';
  readonly append?: string;
}

// Happy DOM keeps XML element namespaces but its namespace query/attribute APIs are incomplete.
// Walk expanded element names and resolve prefixed attributes from the original xmlns bindings.
function descendants(root: Document | Element, name: string): Element[] {
  const result: Element[] = [],
    pending = [...root.children].reverse();
  while (pending.length) {
    const element = pending.pop()!;
    if (element.namespaceURI === W && (name === '*' || element.localName === name))
      result.push(element);
    for (let index = element.children.length - 1; index >= 0; index--)
      pending.push(element.children[index]!);
  }
  return result;
}
function wordAttribute(element: Element, name: string): string | null {
  for (const attribute of element.attributes) {
    const split = attribute.name.split(':');
    if (split.length !== 2 || split[1] !== name) continue;
    let owner: Element | null = element;
    while (owner) {
      const namespace = owner.getAttribute('xmlns:' + split[0]);
      if (namespace !== null) {
        if (namespace === W) return attribute.value;
        break;
      }
      owner = owner.parentElement;
    }
  }
  return null;
}
function tag(control: Element): string | null {
  const properties = [...control.children].find(
    (element) => element.namespaceURI === W && element.localName === 'sdtPr'
  );
  const value = properties ? descendants(properties, 'tag')[0] : null;
  return value ? wordAttribute(value, 'val') : null;
}

/**
 * Independent WordprocessingML result fixture, not a capture from desktop Word.
 * It unwraps retained revision runs and removes discarded runs directly in XML.
 * No production reconciliation or native TreeDocOp decides these fixture results.
 * Desktop Word version/visual interoperability remains V06.
 */
export function externalReviewDocument(
  bytes: Uint8Array,
  choices: Readonly<Record<string, ExternalLabelReview>>
): Uint8Array {
  const entries = unzipSync(bytes),
    xml = new DOMParser().parseFromString(
      strFromU8(entries['word/document.xml']!),
      'application/xml'
    );
  const controls = descendants(xml, 'sdt');
  for (const [id, choice] of Object.entries(choices)) {
    const matches = controls.filter((control) => blockIdFromTag(tag(control) ?? '') === id);
    if (matches.length !== 1) throw new Error('Fixture needs one native owner for ' + id);
    const owner = matches[0]!,
      label = descendants(owner, 'sdt').find((control) => tag(control) === GENERATED_LABEL_TAG),
      paragraphs = descendants(label ?? owner, 'p');
    if (paragraphs.length !== 1) throw new Error('Fixture needs one fallback paragraph for ' + id);
    const paragraph = paragraphs[0]!;
    const revisions = descendants(paragraph, '*').filter(
      (node) =>
        ['ins', 'del'].includes(node.localName) && (!choice.only || node.localName === choice.only)
    );
    for (const revision of revisions.reverse()) {
      if (revision.parentElement?.localName === 'rPr')
        throw new Error('Fixture only models inline revision runs.');
      const keep = (revision.localName === 'ins') === (choice.decision === 'accept');
      if (keep) {
        for (const deleted of descendants(revision, 'delText')) {
          const text = xml.createElementNS(W, deleted.tagName.replace(/delText$/, 't'));
          for (const attribute of deleted.attributes)
            text.setAttribute(attribute.name, attribute.value);
          while (deleted.firstChild) text.append(deleted.firstChild);
          deleted.replaceWith(text);
        }
        while (revision.firstChild)
          revision.parentNode!.insertBefore(revision.firstChild, revision);
      }
      revision.remove();
    }
    if (choice.append) {
      const run = xml.createElementNS(W, 'w:r'),
        text = xml.createElementNS(W, 'w:t');
      text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      text.textContent = choice.append;
      run.append(text);
      paragraph.append(run);
    }
  }
  entries['word/document.xml'] = strToU8(new XMLSerializer().serializeToString(xml));
  return zipSync(entries);
}

export async function externalMixedReviewDocument(decision: 'accept' | 'reject') {
  const editor = propertyEditor(groupFixture());
  try {
    proposeMixedBulk(editor);
    return externalReviewDocument(
      new Uint8Array(await editor.save()),
      Object.fromEntries(['activity', 'visual', 'audio'].map((id) => [id, { decision }]))
    );
  } finally {
    editor.destroy();
  }
}
