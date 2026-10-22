// Independent diagnostic. This imports MIT-licensed Yjs and no editor or evaluated transport.
import * as Y from 'yjs';
import dependency from 'yjs/package.json';

function reproduce(clients: readonly [number, number], receiver: 0 | 1) {
  const docs = [new Y.Doc(), new Y.Doc()];
  docs.forEach((doc, peer) => {
    doc.clientID = clients[peer]!;
  });
  try {
    const alice = docs[0]!.getText('paragraph');
    alice.insert(0, 'Two', { shell: 'original' });
    Y.applyUpdate(docs[1]!, Y.encodeStateAsUpdate(docs[0]!));
    const bob = docs[1]!.getText('paragraph');
    docs[0]!.transact(() => {
      alice.insert(3, 'Changed', { shell: 'replacement', insertion: 'alice', bold: true });
      alice.format(0, 3, { shell: 'deleted', deletion: 'alice', bold: true });
    });
    bob.insert(3, ' Bob', { shell: 'original' });
    const other = 1 - receiver;
    Y.applyUpdate(docs[receiver]!, Y.encodeStateAsUpdate(docs[other]!));
    Y.applyUpdate(docs[other]!, Y.encodeStateAsUpdate(docs[receiver]!));
    const delta = alice.toDelta() as { insert: string; attributes?: Record<string, unknown> }[];
    const proposed = delta
      .filter((span) => !span.attributes?.deletion)
      .map((span) => span.insert)
      .join('');
    return {
      clients,
      receiver,
      delta,
      proposed,
      suffixPreserved: proposed === 'Changed Bob',
      converged: JSON.stringify(alice.toDelta()) === JSON.stringify(bob.toDelta()),
    };
  } finally {
    docs.forEach((doc) => doc.destroy());
  }
}

const cases = (
  [
    [41, 73],
    [73, 41],
  ] as const
).flatMap((clients) => ([0, 1] as const).map((receiver) => reproduce(clients, receiver)));
process.stdout.write(
  JSON.stringify(
    {
      schemaVersion: 1,
      command: 'bun packages/sciobot/test/diagnostics/yjs-format-boundaries.ts',
      dependency: {
        name: dependency.name,
        version: dependency.version,
        license: dependency.license,
      },
      expectedProposed: 'Changed Bob',
      cases,
      finding:
        'Y.Text.format range markers include a concurrent boundary insertion for one client order. Both documents converge, but the unseen suffix inherits deletion and bold. This diagnostic makes no repair and does not establish R29 support.',
    },
    null,
    2
  ) + '\n'
);
