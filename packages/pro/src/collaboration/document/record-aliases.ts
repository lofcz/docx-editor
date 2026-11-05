/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { withPart, type OoxmlNode, type OoxmlPackage } from '@docx-editor.dev/core/store';
import { asLogicalId, type LogicalId } from './identity.ts';
import { expectElement, withChildren } from './node-shapes.ts';
import type { LogicalIdentityMap } from '../document-identity.ts';
import type { DocumentRegistry } from './registry.ts';
import type {
  CanonicalPrimitiveEffect,
  CanonicalPrimitiveJournal,
} from '@docx-editor.dev/core/collaboration/replication';
import { sharedEffect } from './shared-effect.ts';
import { atomicRecordOwnership } from './atomic-record-ownership.ts';

/** Only records owned by an atomic control require eager identity adoption. */
export function atomicRecordAliases(
  identity: LogicalIdentityMap,
  registry: DocumentRegistry,
  journal: CanonicalPrimitiveJournal,
  effects: readonly CanonicalPrimitiveEffect[]
): ReadonlyMap<LogicalId, string> | undefined {
  if (!journal.effects.some((effect) => effect.kind === 'putNode')) return undefined;
  const atomic = atomicRecordOwnership(registry, effects.map(sharedEffect));
  return identity.introducesRecordIdentity(journal, registry, atomic)
    ? identity.recordAliases(registry, atomic)
    : undefined;
}

/** Each retry records only aliases present in the package that it actually installs. */
export function recordAliasProjection(aliases?: ReadonlyMap<LogicalId, string>) {
  let used: ReadonlySet<string> = new Set();
  return {
    package(pkg: OoxmlPackage): OoxmlPackage {
      if (!aliases) return pkg;
      const projected = withRecordAliases(pkg, aliases);
      used = projected.used;
      return projected.package;
    },
    finish(identity: LogicalIdentityMap, changed: boolean): void {
      if (!changed) return;
      if (aliases) identity.retain(used);
      else identity.reset();
    },
  };
}

/** Preserve ordinary creation-result addresses while atomic records adopt shared identities. */
export function withRecordAliases(
  source: OoxmlPackage,
  aliases: ReadonlyMap<LogicalId, string>
): { readonly package: OoxmlPackage; readonly used: ReadonlySet<string> } {
  const used = new Set<string>();
  if (!aliases.size) return { package: source, used };
  // The materializer already bounded and validated this tree. Only local node IDs change;
  // XML attributes, shared identities, payloads, and non-XML parts retain their exact values.
  const visit = (node: OoxmlNode): OoxmlNode => {
    const alias = aliases.get(asLogicalId(node.id));
    if (alias) used.add(alias);
    const id = alias ?? node.id;
    if (node.kind === 'textValue') return id === node.id ? node : { ...node, id };
    const children = node.children.map(visit);
    if (id === node.id && children.every((child, index) => child === node.children[index]))
      return node;
    return withChildren(node, children, id);
  };
  let pkg = source;
  for (const part of source.parts.values()) {
    const root = visit(part.root);
    if (root !== part.root) pkg = withPart(pkg, { ...part, root: expectElement(root) });
  }
  return { package: pkg, used };
}
