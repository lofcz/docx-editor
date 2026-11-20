import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { ROOT, git, json } from './common.mjs';

const INDEX = '.collaboration/extension-archives.json';
const objectId = /^[a-f0-9]{40}$/;

/** Read decisions from their original Git objects after their extension moves to another repository. */
export function archivedDecisions(base) {
  const manifests = existsSync(resolve(ROOT, INDEX)) ? json(INDEX) : [];
  if (git('ls-tree', '--name-only', base, INDEX)) {
    const previous = JSON.parse(git('show', `${base}:${INDEX}`));
    assert.deepEqual(
      manifests.slice(0, previous.length),
      previous,
      'Evidence archives are immutable'
    );
  }
  const archived = new Map();
  for (const manifest of manifests) {
    const { sourceCommit, consumerCommit, decisions, evidence, renamedChangesets } = manifest;
    for (const id of [sourceCommit, consumerCommit, ...decisions, ...evidence])
      assert.match(id, objectId, 'Invalid archived Git object ID');
    git('merge-base', '--is-ancestor', sourceCommit, 'HEAD');
    const tree = new Map(
      git('ls-tree', '-rtz', sourceCommit)
        .split('\0')
        .filter(Boolean)
        .map((line) => {
          const at = line.indexOf('\t');
          const [, type, id] = line.slice(0, at).split(' ');
          return [line.slice(at + 1), { type, id }];
        })
    );
    const evidenceIds = new Set(evidence);
    const decisionIds = new Set(decisions);
    const transferred = new Set();
    for (const [path, { id }] of tree) {
      if (evidenceIds.has(id)) transferred.add(path);
    }
    assert.equal(transferred.size, evidenceIds.size, 'Missing or ambiguous archived test source');
    for (const [file, { id, type }] of tree) {
      if (type !== 'blob' || !file.startsWith('.collaboration/changes/') || !decisionIds.delete(id))
        continue;
      assert.ok(!archived.has(file), `Decision archived twice: ${file}`);
      const raw = git('cat-file', 'blob', id);
      const record = JSON.parse(raw);
      const tests = new Set(
        record.tests.filter((test) => {
          if (tree.get(test)?.type !== 'blob') return false;
          return [...transferred].some((path) => test === path || test.startsWith(`${path}/`));
        })
      );
      const original = tree.get(`.changeset/${record.changeset}.md`)?.id;
      const changeset = original && renamedChangesets[original];
      if (changeset) assert.match(changeset, /^[a-z0-9]+(?:-[a-z0-9]+)*$/);
      archived.set(file, { file, id, record, raw, tests, changeset });
    }
    assert.equal(decisionIds.size, 0, 'Archived decision is absent from its original Git tree');
  }
  return archived;
}

/** Decisions a historical revision had already transferred out of its working tree. */
export function archivedFilesAt(commit, archived) {
  if (!archived.size || !git('ls-tree', '--name-only', commit, INDEX)) return [];
  const manifests = JSON.parse(git('show', `${commit}:${INDEX}`));
  const ids = new Set(manifests.flatMap((manifest) => manifest.decisions));
  return [...archived.values()].filter(({ id }) => ids.has(id)).map(({ file }) => file);
}
