/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  contentControlTextOf,
  WML_NAMESPACE_URI,
  ATOMIC_CONTENT_NAMESPACE,
} from '@docx-editor.dev/core/store';
import { zipDocument } from './document-peer-support';
import {
  destroyReplica,
  joinReplica,
  loadPackage,
  packageFingerprint,
  packageOf,
  seedReplica,
} from './document-support';
import { applyPrimitiveJournal } from '../document/index';
import { asLogicalId } from '../document/identity';

const control = (tag: string, content: string, version?: string) =>
  `<w:sdt ${version ? `xmlns:a="${ATOMIC_CONTENT_NAMESPACE}" a:atomicContent="${version}"` : ''}><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const run = '<w:r><w:t>source</w:t></w:r>';

test('marked inline and display fallbacks seed as records; ordinary controls retain paragraph text ownership', async () => {
  const pkg = loadPackage(
    zipDocument(
      control('atomic-block', `<w:p>${run}</w:p>`, '1') +
        control('ordinary-block', `<w:p>${run}</w:p>`) +
        control('unknown-block', `<w:p>${run}</w:p>`, '2') +
        `<w:p>${control('atomic-inline', run, '1')}${control('ordinary-inline', run)}</w:p>`
    )
  );
  const source = await seedReplica(pkg);
  const peer = joinReplica(source);
  try {
    for (const replica of [source, peer]) {
      const root = packageOf(replica).parts.get(pkg.mainDocumentPart)!.root;
      const controls = new Map(
        contentControlsIn(root).map(({ node }) => [contentControlPropertiesOf(node).tag, node])
      );
      for (const name of ['atomic-block', 'ordinary-block', 'unknown-block']) {
        const content = controls
          .get(name)!
          .children.find((child) => child.kind === 'contentControlContent')!;
        const paragraph = content.children.find((child) => child.kind === 'paragraph')!;
        expect(replica.registry.inline.textOf(paragraph.id) === null).toBe(name === 'atomic-block');
        expect(paragraph.children.some((child) => child.kind === 'run')).toBe(true);
      }
      expect(replica.registry.hasNode(asLogicalId(controls.get('atomic-inline')!.id))).toBe(true);
      expect(replica.registry.hasNode(asLogicalId(controls.get('ordinary-inline')!.id))).toBe(
        false
      );
      expect(packageFingerprint(packageOf(replica))).toBe(packageFingerprint(pkg));
    }
  } finally {
    destroyReplica(source);
    destroyReplica(peer);
  }
});

for (const attachFirst of [true, false])
  test(`wrapping a retained run uses its final owner; removing its last owner deletes it: ${attachFirst}`, async () => {
    const replica = await seedReplica(
      loadPackage(zipDocument(`<w:p>${control('atomic', run, '1')}</w:p>`))
    );
    try {
      const root = packageOf(replica).parts.get('/word/document.xml')!.root;
      const atomic = contentControlsIn(root)[0]!.node;
      const content = atomic.children.find((child) => child.kind === 'contentControlContent')!;
      const retained = content.children.find((child) => child.kind === 'run')!;
      const paragraph = replica.mint.take();
      const attach = {
        kind: 'spliceChildren' as const,
        parentLogicalId: paragraph,
        start: 0,
        deleteCount: 0,
        childLogicalIds: [retained.id],
      };
      const detach = {
        kind: 'spliceChildren' as const,
        parentLogicalId: content.id,
        start: 0,
        deleteCount: 1,
        childLogicalIds: [paragraph],
      };
      expect(
        applyPrimitiveJournal(replica.registry, {
          effects: [
            {
              kind: 'putNode',
              descriptor: {
                logicalId: paragraph,
                kind: 'paragraph',
                qname: { namespaceUri: WML_NAMESPACE_URI, localName: 'p', prefix: 'w' },
              },
            },
            ...(attachFirst ? [attach, detach] : [detach, attach]),
          ],
        }).ok
      ).toBe(true);
      expect(replica.registry.isTombstoned(asLogicalId(retained.id))).toBe(false);
      expect(
        contentControlTextOf(
          contentControlsIn(packageOf(replica).parts.get('/word/document.xml')!.root)[0]!.node
        )
      ).toBe('source');
      const peer = joinReplica(replica);
      try {
        expect(packageFingerprint(packageOf(peer))).toBe(packageFingerprint(packageOf(replica)));
      } finally {
        destroyReplica(peer);
      }
      expect(
        applyPrimitiveJournal(replica.registry, {
          effects: [
            {
              kind: 'spliceChildren',
              parentLogicalId: paragraph,
              start: 0,
              deleteCount: 1,
              childLogicalIds: [],
            },
          ],
        }).ok
      ).toBe(true);
      expect(replica.registry.isTombstoned(asLogicalId(retained.id))).toBe(true);
    } finally {
      destroyReplica(replica);
    }
  });

test('atomic text retains whole replacement entries and malformed fields refuse; ordinary text remains shared characters', async () => {
  const replica = await seedReplica(
    loadPackage(zipDocument(`<w:p>${control('atomic', run, '1')}</w:p>`))
  );
  try {
    const atomic = contentControlsIn(packageOf(replica).parts.get('/word/document.xml')!.root)[0]!
      .node;
    const content = atomic.children.find((child) => child.kind === 'contentControlContent')!;
    const runNode = content.children.find((child) => child.kind === 'run')!;
    const textNode = runNode.children.find((child) => child.kind === 'text')!;
    const value = textNode.children.find((child) => child.kind === 'textValue')!;
    const id = asLogicalId(value.id);
    const record = replica.registry.schema.nodes.get(id)!;
    expect((record.get('atomicText') as Y.Array<string>).toArray()).toEqual(['source']);
    expect(replica.registry.textOf(id).toString()).toBe('');
    replica.registry.spliceText(id, 0, 6, 'replacement');
    expect((record.get('atomicText') as Y.Array<string>).toArray()).toEqual(['replacement']);
    expect(replica.registry.nodeShape(id)?.textLength).toBe(11);
    record.set('atomicText', { invalid: true });
    expect(replica.registry.record(id)).toBeNull();
    expect(replica.registry.nodeShape(id)).toBeNull();
    const ordinary = replica.mint.take();
    replica.registry.putText(ordinary, 'ordinary');
    replica.registry.spliceText(ordinary, 0, 0, 'editable ');
    expect(replica.registry.textOf(ordinary).toString()).toBe('editable ordinary');
    expect(replica.registry.schema.nodes.get(ordinary)!.has('atomicText')).toBe(false);
  } finally {
    destroyReplica(replica);
  }
});
