import {
  applyTreeOp,
  directParagraphProperties,
  ensureHyperlinkRelationship,
  escapeXml,
  readOoxmlPackage,
  runPropertyEdits,
  withPart,
  type OoxmlPackage,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { docxFixture } from './document';
import { protectedDocument } from './protected-document';
import { projectSourcePackage, sourceControl, sourceRevision } from './target-source-fixture';
import { buildBlockIndex } from '../../src/blocks/block-index';
import { withPayloadRecord } from '../../src/blocks/payload-opc-codec';
import type { BlockPayload } from '../../src/blocks/payload-schema';
import type { TargetSourceInput } from '../../src/diff/target-source';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import { directParagraphTextOps } from '../../src/diff/paragraph-text';

export const targetParagraph = (text: string) =>
  `<w:p><w:pPr><keep:extra xmlns:keep="urn:keep" keep:flag="paragraph"/></w:pPr><w:r><w:rPr><keep:extra xmlns:keep="urn:keep" keep:flag="run"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
export async function targetDocumentInput(
  body: string,
  relationships?: string
): Promise<TargetSourceInput> {
  const loaded = readOoxmlPackage(
    docxFixture(
      body + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
      relationships ? { relationships } : {}
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return {
    revision: sourceRevision,
    pkg: loaded.package,
    projection: await projectSourcePackage(loaded.package),
  };
}
export async function targetLessonInput(): Promise<TargetSourceInput> {
  const body =
    sourceControl('title', targetParagraph('Lesson')) +
    sourceControl(
      'phase',
      sourceControl(
        'activity',
        targetParagraph('First question.') +
          `<w:p>${sourceControl('source', '<w:r><w:t>Source</w:t></w:r>')}</w:p>` +
          targetParagraph('Last question.')
      )
    );
  const base = await targetDocumentInput(body);
  let pkg = base.pkg;
  for (const [id, type, attributes] of [
    ['title', 'lesson_title', { grade: 7 }],
    ['phase', 'lesson_phase', { phaseType: 'evocation' }],
    [
      'activity',
      'lesson_activity',
      { name: 'Warmup', duration: 10, privateData: { preserve: true } },
    ],
    [
      'source',
      'source',
      {
        title: 'Source',
        url: 'https://example.com',
        trust: 'verified',
        privateData: { preserve: true },
      },
    ],
  ] as const) {
    const stored = withPayloadRecord(pkg, { id, type, attributes, version: 1 }, id);
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  return { revision: sourceRevision, pkg, projection: await projectSourcePackage(pkg) };
}

/** Retained-operation helper for immutable test oracles. Live execution belongs to A01/R05. */
export function applyTargetRetained(input: OoxmlPackage, plan: CanonicalEditPlan): OoxmlPackage {
  let pkg = input;
  for (const edit of plan.operations) {
    let part = pkg.parts.get(edit.partName)!;
    const apply = (operation: TreeDocOp) => {
      const result = applyTreeOp(part, operation);
      if (!result.ok) throw new Error(result.reason);
      part = result.part;
    };
    if (edit.kind === 'text') directParagraphTextOps(edit).forEach(apply);
    else if (edit.kind === 'format' && edit.target.kind === 'run') {
      const target = edit.target,
        removed = new Set(
          edit.changes.filter((change) => !change.after.present).map((change) => change.name)
        ),
        properties = edit.changes.flatMap((change) =>
          change.after.present
            ? [{ localName: change.name, attributes: change.after.value as Record<string, string> }]
            : []
        );
      const native = runPropertyEdits(
        part,
        target.paragraphId,
        target.start,
        target.end,
        properties,
        'proposed'
      );
      for (const change of native)
        apply({
          op: 'setRunProperties',
          paragraphId: target.paragraphId,
          start: change.start,
          end: change.end,
          properties: change.properties.filter((property) => !removed.has(property.localName)),
        });
    } else if (edit.kind === 'format' && edit.target.kind === 'paragraph') {
      const properties = new Map(
        directParagraphProperties(part, edit.target.paragraphId).map((property) => [
          property.localName,
          property,
        ])
      );
      for (const change of edit.changes) {
        if (!change.after.present) properties.delete(change.name);
        else
          properties.set(change.name, {
            localName: change.name,
            attributes: change.after.value as Record<string, string>,
          });
      }
      apply({
        op: 'setParagraphProperties',
        paragraphId: edit.target.paragraphId,
        properties: [...properties.values()],
      });
    } else if (edit.kind === 'format' && edit.target.kind === 'hyperlink') {
      const change = edit.changes.find((change) => change.name === 'target');
      if (!change?.after.present) throw new Error('Unsupported hyperlink fixture');
      const value = change.after.value as { kind: string; authored: string };
      const link = ensureHyperlinkRelationship(pkg, value.authored, edit.partName);
      if (!link) throw new Error('Unsafe fixture target');
      pkg = link.pkg;
      apply({
        op: 'setHyperlinkTarget',
        linkId: edit.target.hyperlinkId,
        relationshipId: link.relationshipId,
      });
    } else if (edit.kind === 'payload') {
      const current = buildBlockIndex(pkg).byId.get(edit.blockId)!;
      let payload = JSON.parse(JSON.stringify(current.payload)) as BlockPayload;
      if (!edit.path.length && edit.after.present)
        payload = edit.after.value as unknown as BlockPayload;
      else if (
        edit.path.length === 2 &&
        edit.path[0] === 'attributes' &&
        typeof edit.path[1] === 'string'
      ) {
        if (edit.after.present)
          Object.defineProperty(payload.attributes, edit.path[1], {
            value: edit.after.value,
            enumerable: true,
            writable: true,
            configurable: true,
          });
        else delete (payload.attributes as Record<string, unknown>)[edit.path[1]];
      } else throw new Error('Unsupported fixture payload path');
      const stored = withPayloadRecord(pkg, payload, edit.blockId);
      if (!stored.ok) throw new Error(stored.reason);
      pkg = stored.pkg;
    } else throw new Error(`Unsupported retained fixture operation: ${edit.kind}`);
    if (edit.kind !== 'payload') pkg = withPart(pkg, part);
  }
  return pkg;
}

export interface ProtectedTargetOptions {
  readonly relationships?: string;
  readonly numbering?: string;
  readonly payloads?: readonly BlockPayload[];
}
/** Embed a C01 body in imported Word content with a header, footer, notes, media, and unknown parts. */
export async function protectedTargetInput(
  body: string,
  options: ProtectedTargetOptions = {}
): Promise<TargetSourceInput> {
  const entries = unzipSync(protectedDocument()),
    document = strFromU8(entries['word/document.xml']!),
    bodyStart = document.indexOf('<w:body>') + '<w:body>'.length,
    sectionStart = document.lastIndexOf('<w:sectPr>'),
    relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  entries['word/document.xml'] = strToU8(
    document.slice(0, bodyStart) + body + document.slice(sectionStart)
  );
  const extra =
    (options.relationships ?? '') +
    (options.numbering
      ? `<Relationship Id="numbering" Type="${relationship}/numbering" Target="numbering.xml"/>`
      : '');
  entries['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(entries['word/_rels/document.xml.rels']!).replace(
      '</Relationships>',
      `${extra}</Relationships>`
    )
  );
  if (options.numbering) {
    entries['word/numbering.xml'] = strToU8(options.numbering);
    entries['[Content_Types].xml'] = strToU8(
      strFromU8(entries['[Content_Types].xml']!).replace(
        '</Types>',
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
      )
    );
  }
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  let pkg = loaded.package;
  for (const payload of options.payloads ?? []) {
    const stored = withPayloadRecord(pkg, payload, payload.id);
    if (!stored.ok) throw new Error(stored.reason);
    pkg = stored.pkg;
  }
  return { revision: sourceRevision, pkg, projection: await projectSourcePackage(pkg) };
}
