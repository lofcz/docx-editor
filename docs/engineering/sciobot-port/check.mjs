import { readFileSync, readdirSync, existsSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const directory = dirname(fileURLToPath(import.meta.url));
const root = resolve(directory, '../../..');
const index = resolve(root, 'SCIOBOT_PORT_TODO.md');
const files = readdirSync(directory)
  .filter((name) => name.endsWith('.md'))
  .map((name) => resolve(directory, name));
const errors = [];
const tasks = new Map();
const documents = new Map([index, ...files].map((file) => [file, readFileSync(file, 'utf8')]));
let productionAudit = '';
try {
  productionAudit = execFileSync(
    'bun',
    [resolve(root, 'scripts/sciobot-port/check-production-coverage.mjs')],
    {
      cwd: root,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    }
  );
} catch (error) {
  errors.push(`Production coverage audit failed: ${String(error.stderr ?? error.message).trim()}`);
}

for (const file of files) {
  const markdown = documents.get(file);
  const headings = [...markdown.matchAll(/^## ([A-Z]\d+)\s*$/gm)];
  for (let position = 0; position < headings.length; position++) {
    const heading = headings[position];
    const id = heading[1];
    if (id.startsWith('S')) continue;
    const section = markdown.slice(heading.index, headings[position + 1]?.index ?? markdown.length);
    const checkboxes = [...section.matchAll(/^- \[([ x])\] (.+)$/gm)];
    const dependencies = section.match(/\*\*Depends:\*\* ([^\n]+)/)?.[1];
    if (tasks.has(id)) errors.push(`Duplicate leaf: ${id}`);
    if (checkboxes.length !== 1) errors.push(`${id}: expected one checkbox`);
    if (!dependencies) errors.push(`${id}: missing dependencies`);
    for (const field of [
      'Primary files',
      'Read',
      'Change',
      'Close when',
      'Check',
      'Claim',
      'Evidence',
    ]) {
      if (!section.includes(`**${field}:**`)) errors.push(`${id}: missing ${field}`);
    }
    tasks.set(id, {
      id,
      title: checkboxes[0]?.[2] ?? id,
      complete: checkboxes[0]?.[1] === 'x',
      dependencies: [...(dependencies ?? '').matchAll(/\[([A-Z]\d+)\]/g)].map((match) => match[1]),
      allOthers: dependencies === 'Every other leaf in this backlog.',
      evidence: section.split('**Evidence:**')[1]?.trim() ?? '',
      claim: section.match(/\*\*Claim:\*\* (.*?)\s+\*\*Evidence:/s)?.[1]?.trim() ?? '',
    });
  }
}

for (const task of tasks.values()) {
  if (task.allOthers) task.dependencies = [...tasks.keys()].filter((id) => id !== task.id);
  for (const dependency of task.dependencies) {
    if (!tasks.has(dependency)) errors.push(`${task.id}: unknown dependency ${dependency}`);
    if (task.complete && !tasks.get(dependency)?.complete) {
      errors.push(`${task.id}: closed before dependency ${dependency}`);
    }
  }
  if (task.complete && (!task.evidence || /^pending\.?$/i.test(task.evidence))) {
    errors.push(`${task.id}: closed without evidence`);
  }
  if (task.complete && (!task.claim || /^unclaimed\.?$/i.test(task.claim))) {
    errors.push(`${task.id}: closed without an owner`);
  }
}

const active = new Set();
const visited = new Set();
function visit(id) {
  if (active.has(id)) {
    errors.push(`Dependency cycle at ${id}`);
    return;
  }
  if (visited.has(id)) return;
  active.add(id);
  for (const dependency of tasks.get(id)?.dependencies ?? []) visit(dependency);
  active.delete(id);
  visited.add(id);
}
for (const id of tasks.keys()) visit(id);

// Verify local links, including source files and stable leaf anchors.
for (const [file, markdown] of documents) {
  for (const match of markdown.matchAll(/\[[^\]\n]+\]\(([^)\n]+)\)/g)) {
    const target = match[1];
    if (/^[a-z]+:/i.test(target)) continue;
    const [path, anchor] = target.split('#');
    const resolved = path ? resolve(dirname(file), path) : file;
    if (!existsSync(resolved)) {
      errors.push(`${file}: missing link ${target}`);
      continue;
    }
    if (anchor && statSync(resolved).isFile()) {
      const body = documents.get(resolved) ?? readFileSync(resolved, 'utf8');
      const anchors = [...body.matchAll(/^#{1,6} (.+)$/gm)].map((heading) =>
        heading[1]
          .toLowerCase()
          .replace(/[^\w\s-]/g, '')
          .replace(/\s+/g, '-')
      );
      if (!anchors.includes(anchor)) errors.push(`${file}: missing anchor ${target}`);
    }
  }
}

if (errors.length) {
  process.stderr.write(`${errors.join('\n')}\n`);
  process.exitCode = 1;
} else {
  process.stdout.write(productionAudit);
  const completed = [...tasks.values()].filter((task) => task.complete).length;
  const ready = [...tasks.values()].filter(
    (task) => !task.complete && task.dependencies.every((id) => tasks.get(id)?.complete)
  );
  process.stdout.write(
    `${tasks.size} leaves; ${completed} closed; ${ready.length} eligible by dependency.\n`
  );
  process.stdout.write('Links, source paths, dependencies, and closure fields pass validation.\n');
  for (const task of ready) process.stdout.write(`${task.id}: ${task.title}\n`);
}
