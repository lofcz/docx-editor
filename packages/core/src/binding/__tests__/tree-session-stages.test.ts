import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { openTreeSession, type TreeDocxSession } from '../tree-session.ts';
import { paragraphTextOf } from '../../store/store/tree-ops.ts';
import { allParagraphs } from '../tree-binding.ts';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  ensureHyperlinkRelationship,
  type TreeTransactionContext,
} from '../../store/index.ts';
import type { TreeApplyOptions } from '../tree-session.ts';
import { ACTOR_ID_STRIPE, actorStripe } from '../../store/package/actor-scoped-ids.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function open(
  body = p('Alpha') + p('Tail'),
  extra: { parts?: Record<string, string>; relationships?: string; overrides?: string } = {}
): TreeDocxSession {
  const bytes = zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${extra.overrides ?? ''}</Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`
    ),
    ...(extra.relationships
      ? {
          'word/_rels/document.xml.rels': strToU8(
            `<Relationships xmlns="${REL}">${extra.relationships}</Relationships>`
          ),
        }
      : {}),
    ...Object.fromEntries(
      Object.entries(extra.parts ?? {}).map(([name, xml]) => [name, strToU8(xml)])
    ),
  });
  const loaded = openTreeSession(bytes);
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.session;
}

test('a native stage reads earlier operations and publishes one complete result', () => {
  const session = open(),
    paragraphId = session.paragraphIds()[0]!,
    original = session.currentPackage();
  let notifications = 0;
  session.subscribe(() => {
    notifications++;
    expect(session.bodyText()).toBe('Alpha Beta Gamma\nTail');
  });
  const result = session.applyTreeOps(
    [{ op: 'insertText', paragraphId, offset: 5, text: ' Beta' }],
    undefined,
    undefined,
    { kind: 'body' },
    {
      stage: (ctx) => {
        const working = ctx.currentPackage();
        expect(session.currentPackage()).toBe(original);
        expect(notifications).toBe(0);
        const part = working.parts.get(working.mainDocumentPart)!;
        expect(paragraphTextOf(part, paragraphId)).toBe('Alpha Beta');
        expect(ctx.apply({ op: 'insertText', paragraphId, offset: 10, text: ' Gamma' })).toBe(true);
        expect(
          paragraphTextOf(ctx.currentPackage().parts.get(working.mainDocumentPart)!, paragraphId)
        ).toBe('Alpha Beta Gamma');
      },
    }
  );
  expect(result).toMatchObject({ committed: true, rejected: false, opCount: 2 });
  expect(notifications).toBe(1);
  expect(session.packageRevision()).toBe(1);
  const reopened = openTreeSession(session.save());
  expect(reopened.ok && reopened.session.bodyText()).toBe('Alpha Beta Gamma\nTail');
});

const resource = (ctx: TreeTransactionContext) => {
  expect(
    ctx.applyPackage((pkg) => {
      const made = ensureHyperlinkRelationship(
        pkg,
        'https://example.org/stage',
        pkg.mainDocumentPart
      );
      if (!made) throw new Error('Cannot make test resource.');
      return made.pkg;
    })
  ).toBe(true);
};
const fingerprint = (session: TreeDocxSession) =>
  [...session.currentPackage().parts].map(([name, part]) => [
    name,
    canonicalOoxmlFingerprint(part),
  ]);

for (const atomic of [false, true]) {
  const apply = (session: TreeDocxSession, options: TreeApplyOptions) =>
    atomic
      ? session.applyTreeOpsAtomic([], options)
      : session.applyTreeOps([], undefined, undefined, { kind: 'body' }, options);

  test(`stage metadata and native edits have one publication and history unit (atomic=${atomic})`, () => {
    const session = open(),
      paragraphId = session.paragraphIds()[0]!,
      original = fingerprint(session);
    let count = 0;
    const unsubscribe = session.subscribe(() => count++);
    const order: string[] = [];
    const result = apply(session, {
      packageEdits: [
        (pkg) => {
          order.push('before');
          return pkg;
        },
      ],
      stage: (ctx) => {
        order.push('stage');
        resource(ctx);
        expect(ctx.currentPackage().parts.has('/word/_rels/document.xml.rels')).toBe(true);
        expect(ctx.apply({ op: 'insertText', paragraphId, offset: 5, text: ' staged' })).toBe(true);
        expect(session.bodyText()).toBe('Alpha\nTail');
      },
      packageEditsAfterOps: [
        (pkg) => {
          order.push('after');
          expect(paragraphTextOf(pkg.parts.get(pkg.mainDocumentPart)!, paragraphId)).toBe(
            'Alpha staged'
          );
          return pkg;
        },
      ],
    });
    expect(result).toMatchObject({ committed: true, rejected: false, opCount: 1 });
    expect(order).toEqual(['before', 'stage', 'after']);
    expect(count).toBe(1);
    unsubscribe();
    const complete = fingerprint(session);
    session.undo();
    expect(fingerprint(session)).toEqual(original);
    expect(session.canUndo()).toBe(false);
    session.redo();
    expect(fingerprint(session)).toEqual(complete);
    expect(session.canRedo()).toBe(false);
  });

  for (const thrown of [false, true]) {
    test(`failed stages preserve canonical identity, metadata, notifications, and prior redo (atomic=${atomic}, throw=${thrown})`, () => {
      const session = open(),
        paragraphId = session.paragraphIds()[0]!;
      session.applyTreeOps([{ op: 'insertText', paragraphId, offset: 5, text: ' redo' }]);
      session.undo();
      const original = session.currentPackage(),
        revision = session.packageRevision(),
        before = fingerprint(session);
      let notifications = 0,
        finalHooks = 0;
      session.subscribe(() => notifications++);
      const run = () =>
        apply(session, {
          stage: (ctx) => {
            resource(ctx);
            expect(
              ctx.apply({ op: 'insertText', paragraphId, offset: 5, text: ' transient' })
            ).toBe(true);
            if (thrown) throw new Error('Injected stage failure.');
            expect(
              ctx.apply({ op: 'insertText', paragraphId, offset: 1000, text: 'invalid' })
            ).toBe(false);
            expect(ctx.apply({ op: 'insertText', paragraphId, offset: 0, text: 'ignored' })).toBe(
              false
            );
          },
          packageEditsAfterOps: [
            (pkg) => {
              finalHooks++;
              return pkg;
            },
          ],
        });
      if (thrown) expect(run).toThrow('Injected stage failure.');
      else expect(run()).toMatchObject({ committed: false, rejected: true });
      expect(session.currentPackage()).toBe(original);
      expect(fingerprint(session)).toEqual(before);
      expect(session.packageRevision()).toBe(revision);
      expect(notifications).toBe(0);
      expect(finalHooks).toBe(0);
      expect(session.canUndo()).toBe(false);
      expect(session.canRedo()).toBe(true);
      session.redo();
      expect(session.bodyText()).toBe('Alpha redo\nTail');
      expect(session.currentPackage().parts.has('/word/_rels/document.xml.rels')).toBe(false);
    });
  }

  test(`an empty stage is a no-op and retains prior redo (atomic=${atomic})`, () => {
    const session = open(),
      paragraphId = session.paragraphIds()[0]!;
    session.applyTreeOps([{ op: 'insertText', paragraphId, offset: 5, text: ' redo' }]);
    session.undo();
    const original = session.currentPackage(),
      revision = session.packageRevision();
    let notifications = 0;
    session.subscribe(() => notifications++);
    expect(
      apply(session, {
        stage: (ctx) => {
          expect(ctx.currentPackage()).toBe(original);
        },
      })
    ).toMatchObject({ committed: false, rejected: false, opCount: 0 });
    expect(session.currentPackage()).toBe(original);
    expect(session.packageRevision()).toBe(revision);
    expect(notifications).toBe(0);
    expect(session.canRedo()).toBe(true);
  });

  test(`a retained context cannot write after its stage returns (atomic=${atomic})`, () => {
    const session = open();
    let retained: TreeTransactionContext | undefined;
    apply(session, {
      stage: (ctx) => {
        retained = ctx;
      },
    });
    expect(() => retained!.currentPackage()).toThrow('no longer active');
    expect(() =>
      retained!.apply({
        op: 'insertText',
        paragraphId: session.paragraphIds()[0]!,
        offset: 0,
        text: 'late',
      })
    ).toThrow('no longer active');
    expect(() => retained!.applyPackage((pkg) => pkg)).toThrow('no longer active');
    expect(() => retained!.selectionAfter(null)).toThrow('no longer active');
    expect(session.bodyText()).toBe('Alpha\nTail');
    expect(session.packageRevision()).toBe(0);
  });

  test(`asynchronous stages roll back before their continuation runs (atomic=${atomic})`, async () => {
    const session = open(),
      paragraphId = session.paragraphIds()[0]!,
      original = session.currentPackage();
    let lateError = '';
    expect(() =>
      apply(session, {
        stage: async (ctx) => {
          ctx.apply({ op: 'insertText', paragraphId, offset: 5, text: ' transient' });
          await Promise.resolve();
          try {
            ctx.apply({ op: 'insertText', paragraphId, offset: 0, text: 'late' });
          } catch (error) {
            lateError = (error as Error).message;
          }
        },
      })
    ).toThrow('finish synchronously');
    await Promise.resolve();
    expect(lateError).toContain('no longer active');
    expect(session.currentPackage()).toBe(original);
    expect(session.packageRevision()).toBe(0);
    expect(session.canUndo()).toBe(false);
  });

  test(`caught lifecycle violations still abort a stage (atomic=${atomic})`, () => {
    const session = open(),
      original = session.currentPackage();
    expect(() =>
      apply(session, {
        stage: (ctx) => {
          ctx.apply({
            op: 'insertText',
            paragraphId: session.paragraphIds()[0]!,
            offset: 5,
            text: ' transient',
          });
          try {
            ctx.apply({
              op: 'insertNote',
              noteKind: 'footnote',
              paragraphId: session.paragraphIds()[0]!,
              offset: 0,
            });
          } catch {
            /* A caller cannot convert a prohibited stage into a successful partial commit. */
          }
        },
      })
    ).toThrow('Lifecycle operations');
    expect(session.currentPackage()).toBe(original);
    expect(session.packageRevision()).toBe(0);
    expect(session.canUndo()).toBe(false);
  });
}

test.each([false, true])(
  'a later stage retains native control locks (cross-part=%s)',
  (crossPart) => {
    const session = open(
      p('Alpha') +
        '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
        p('Locked') +
        '</w:sdtContent></w:sdt>'
    );
    const original = session.currentPackage(),
      [free, locked] = session.paragraphIds();
    const result = session.applyTreeOps(
      [{ op: 'insertText', paragraphId: free!, offset: 5, text: ' transient' }],
      undefined,
      undefined,
      { kind: 'body' },
      {
        stage: (ctx) => {
          const op = {
            op: 'insertText' as const,
            paragraphId: locked!,
            offset: 0,
            text: 'invalid',
          };
          expect(crossPart ? ctx.applyTo(original.mainDocumentPart, op) : ctx.apply(op)).toBe(
            false
          );
        },
      }
    );
    expect(result).toMatchObject({ committed: false, rejected: true });
    expect(session.currentPackage()).toBe(original);
    expect(session.packageRevision()).toBe(0);
  }
);

test('staged revision allocation retains the transaction actor', () => {
  const session = open(
      '<w:p><w:ins w:id="17" w:author="Bot"><w:r><w:t>Old</w:t></w:r></w:ins></w:p>' + p('Tail')
    ),
    paragraphId = session.paragraphIds()[0]!;
  const result = session.applyTreeOps(
    [{ op: 'rejectAllRevisions' }],
    undefined,
    undefined,
    { kind: 'body' },
    {
      actorId: 'alice',
      stage: (ctx) => {
        expect(
          paragraphTextOf(ctx.currentPackage().parts.get('/word/document.xml')!, paragraphId)
        ).toBe('');
        expect(
          ctx.apply({
            op: 'insertText',
            paragraphId,
            offset: 0,
            text: 'New',
            revision: { author: 'Bot', date: '2026-10-21T09:00:00Z' },
          })
        ).toBe(true);
      },
    }
  );
  expect(result.committed).toBe(true);
  const sites = collectRevisionSites(session.part());
  expect(sites).toHaveLength(1);
  const id = sites[0]!.node.attributes.find((a) => a.localName === 'id')?.value;
  expect(id).not.toBe('17');
  expect(Number(id) % ACTOR_ID_STRIPE).toBe(actorStripe('alice'));
});

test.each([false, true])(
  'staged note deletion uses the intermediate package cleanup (cross-part=%s)',
  (crossPart) => {
    const session = open(
      '<w:p><w:r><w:t>Alpha</w:t><w:footnoteReference w:id="1"/></w:r></w:p>' + p('Tail'),
      {
        relationships: `<Relationship Id="rId7" Type="${R}/footnotes" Target="footnotes.xml"/>`,
        overrides:
          '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
        parts: {
          'word/footnotes.xml': `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${p('Note body')}</w:footnote></w:footnotes>`,
        },
      }
    );
    const original = fingerprint(session),
      paragraphId = session.paragraphIds()[0]!;
    const result = session.applyTreeOps(
      [{ op: 'insertText', paragraphId, offset: 0, text: 'prefix ' }],
      undefined,
      undefined,
      { kind: 'body' },
      {
        stage: (ctx) => {
          expect(
            paragraphTextOf(ctx.currentPackage().parts.get('/word/document.xml')!, paragraphId)
          ).toHaveLength(13);
          const op = { op: 'deleteText' as const, paragraphId, start: 12, end: 13 };
          expect(crossPart ? ctx.applyTo('/word/document.xml', op) : ctx.apply(op)).toBe(true);
        },
      }
    );
    expect(result.committed).toBe(true);
    const notes = session.currentPackage().parts.get('/word/footnotes.xml');
    expect(
      notes ? allParagraphs(notes).map((p) => paragraphTextOf(notes, p.id)) : []
    ).not.toContain('Note body');
    const complete = fingerprint(session);
    session.undo();
    expect(fingerprint(session)).toEqual(original);
    session.redo();
    expect(fingerprint(session)).toEqual(complete);
  }
);

for (const failure of [false, true]) {
  test(`dependent stages span body/header parts without partial publication (failure=${failure})`, () => {
    const session = open(
      p('Alpha') + `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/></w:sectPr>`,
      {
        relationships: `<Relationship Id="rId7" Type="${R}/header" Target="header1.xml"/>`,
        overrides:
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
        parts: { 'word/header1.xml': `<w:hdr xmlns:w="${W}">${p('Header')}</w:hdr>` },
      }
    );
    // Opening an imported story supplies its native paragraph identities before the baseline.
    expect(session.storyText({ kind: 'headerFooter', rId: 'rId7' })).toBe('Header');
    const paragraphId = session.paragraphIds()[0]!,
      original = session.currentPackage(),
      before = fingerprint(session),
      revision = session.packageRevision();
    let notifications = 0;
    session.subscribe(() => notifications++);
    const result = session.applyTreeOpsAtomic(
      [
        {
          scope: { kind: 'body' },
          ops: [{ op: 'insertText', paragraphId, offset: 5, text: ' staged' }],
        },
      ],
      {
        stage: (ctx) => {
          const header = ctx.currentPackage().parts.get('/word/header1.xml')!,
            id = allParagraphs(header)[0]!.id;
          expect(
            paragraphTextOf(ctx.currentPackage().parts.get('/word/document.xml')!, paragraphId)
          ).toBe('Alpha staged');
          expect(
            ctx.applyTo(header.name, {
              op: 'insertText',
              paragraphId: id,
              offset: 6,
              text: ' staged',
            })
          ).toBe(true);
          resource(ctx);
          if (failure)
            expect(
              ctx.applyTo('/word/missing.xml', {
                op: 'insertText',
                paragraphId: id,
                offset: 0,
                text: 'invalid',
              })
            ).toBe(false);
        },
      }
    );
    if (failure) {
      expect(result).toMatchObject({ committed: false, rejected: true, reason: 'unknown-part' });
      expect(session.currentPackage()).toBe(original);
      expect(fingerprint(session)).toEqual(before);
      expect(session.packageRevision()).toBe(revision);
      expect(notifications).toBe(0);
      expect(session.canUndo()).toBe(false);
    } else {
      expect(result).toMatchObject({ committed: true, rejected: false, opCount: 2 });
      expect(notifications).toBe(1);
      expect(session.bodyText()).toBe('Alpha staged');
      expect(session.storyText({ kind: 'headerFooter', rId: 'rId7' })).toBe('Header staged');
      const complete = fingerprint(session);
      session.undo();
      expect(fingerprint(session)).toEqual(before);
      expect(session.canUndo()).toBe(false);
      session.redo();
      expect(fingerprint(session)).toEqual(complete);
      const reopened = openTreeSession(session.save());
      expect(reopened.ok && reopened.session.bodyText()).toBe('Alpha staged');
      expect(reopened.ok && reopened.session.storyText({ kind: 'headerFooter', rId: 'rId7' })).toBe(
        'Header staged'
      );
    }
  });
}

test.each([false, true])(
  'staged document protection refuses all writes (metadata=%s)',
  (metadata) => {
    const session = open(p('Alpha'), {
      relationships: `<Relationship Id="rId7" Type="${R}/settings" Target="settings.xml"/>`,
      overrides:
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
      parts: {
        'word/settings.xml': `<w:settings xmlns:w="${W}"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>`,
      },
    });
    const original = session.currentPackage();
    const result = session.applyTreeOps(
      [],
      undefined,
      undefined,
      { kind: 'body' },
      {
        stage: (ctx) => {
          const accepted = metadata
            ? ctx.applyPackage(
                (pkg) =>
                  ensureHyperlinkRelationship(
                    pkg,
                    'https://example.org/protected',
                    pkg.mainDocumentPart
                  )!.pkg
              )
            : ctx.apply({
                op: 'insertText',
                paragraphId: session.paragraphIds()[0]!,
                offset: 0,
                text: 'invalid',
              });
          expect(accepted).toBe(false);
        },
      }
    );
    expect(result).toMatchObject({ committed: false, rejected: true });
    expect(session.currentPackage()).toBe(original);
    expect(session.packageRevision()).toBe(0);
    expect(session.canUndo()).toBe(false);
  }
);

test('a refused initial operation does not call the dependent stage', () => {
  const session = open(),
    original = session.currentPackage();
  let called = false;
  const result = session.applyTreeOps(
    [{ op: 'insertText', paragraphId: session.paragraphIds()[0]!, offset: 1000, text: 'invalid' }],
    undefined,
    undefined,
    { kind: 'body' },
    {
      stage: () => {
        called = true;
      },
    }
  );
  expect(result).toMatchObject({ committed: false, rejected: true });
  expect(called).toBe(false);
  expect(session.currentPackage()).toBe(original);
  expect(session.packageRevision()).toBe(0);
});

test('a stage cannot accompany an initial lifecycle operation', () => {
  const session = open(),
    original = session.currentPackage();
  let called = false;
  const result = session.applyTreeOps(
    [
      {
        op: 'insertNote',
        noteKind: 'footnote',
        paragraphId: session.paragraphIds()[0]!,
        offset: 0,
      },
    ],
    undefined,
    undefined,
    { kind: 'body' },
    {
      stage: () => {
        called = true;
      },
    }
  );
  expect(result).toMatchObject({ committed: false, rejected: true, reason: 'invalidArgs' });
  expect(called).toBe(false);
  expect(session.currentPackage()).toBe(original);
  expect(session.packageRevision()).toBe(0);
});

test('a stage-only native transaction applies operations and has one undo unit', () => {
  const session = open(),
    paragraphId = session.paragraphIds()[0]!;
  const result = session.applyTreeOps(
    [],
    undefined,
    undefined,
    { kind: 'body' },
    {
      stage: (ctx) => {
        expect(ctx.apply({ op: 'insertText', paragraphId, offset: 5, text: ' staged' })).toBe(true);
      },
    }
  );
  expect(result).toMatchObject({ committed: true, rejected: false, opCount: 1 });
  expect(session.bodyText()).toBe('Alpha staged\nTail');
  session.undo();
  expect(session.bodyText()).toBe('Alpha\nTail');
  expect(session.canUndo()).toBe(false);
  session.redo();
  expect(session.bodyText()).toBe('Alpha staged\nTail');
  expect(session.canRedo()).toBe(false);
});
