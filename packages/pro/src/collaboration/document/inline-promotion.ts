/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { OoxmlNode } from '@docx-editor.dev/core/store';
import { idOf, type LogicalId } from './identity.ts';
import type { SharedEffect } from './shared-effect.ts';

/** Materialize an inline shell as records before a canonical move takes it outside a paragraph. */
export function promoteInlineTree(
  node: OoxmlNode,
  known: (id: LogicalId) => boolean
): SharedEffect[] {
  const effects: SharedEffect[] = [];
  const visit = (node: OoxmlNode) => {
    const id = idOf(node);
    if (known(id)) return;
    if (node.kind === 'textValue') {
      effects.push({ kind: 'putNode', descriptor: { logicalId: id, kind: 'textValue' } });
      if (node.value)
        effects.push({
          kind: 'spliceText',
          logicalId: id,
          utf16Start: 0,
          deleteCount: 0,
          insert: node.value,
        });
      return;
    }
    effects.push({
      kind: 'putNode',
      descriptor: {
        logicalId: id,
        kind: node.kind,
        qname: {
          namespaceUri: node.namespaceUri,
          localName: node.localName,
          ...(node.prefix === undefined ? {} : { prefix: node.prefix }),
        },
      },
    });
    for (const attribute of node.attributes)
      effects.push({
        kind: 'setAttribute',
        logicalId: id,
        qname: {
          namespaceUri: attribute.namespaceUri,
          localName: attribute.localName,
          ...(attribute.prefix === undefined ? {} : { prefix: attribute.prefix }),
        },
        value: attribute.value,
      });
    for (const binding of node.namespaceBindings)
      effects.push({
        kind: 'setNamespaceBinding',
        logicalId: id,
        prefix: binding.prefix,
        uri: binding.namespaceUri,
      });
    for (const child of node.children) visit(child);
    if (node.children.length)
      effects.push({
        kind: 'spliceChildren',
        parentLogicalId: id,
        start: 0,
        deleteCount: 0,
        childLogicalIds: node.children.map(idOf),
      });
  };
  visit(node);
  return effects;
}
