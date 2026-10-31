import {
  applyTreeOp,
  bodyStoryRoot,
  contentControlPropertiesOf,
  directParagraphProperties,
  ensureHyperlinkRelationship,
  ensureListDefinition,
  escapeXml,
  findNode,
  insertChildren,
  readOoxmlPart,
  removeNode,
  withPart,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import type { CanonicalEditPlan } from '../../src/diff/plan';
import type { ListTarget } from '../../src/diff/list-types';
import type { TargetAstNode } from '../../src/projection/parse-target';
import { checklistTag } from '../../src/registry/adapter-names';
import { W } from './document';
import { applyTargetRetained } from './target-document';
import { materializeInsertedRow } from './inserted-row-oracle';
import { materializeInsertedColumn } from './inserted-column-oracle';

let epoch = 0;
/** Immutable native acceptance oracle, not a live executor. It materializes only new target paragraphs. */
export function applyTargetPlanOracle(input: OoxmlPackage, plan: CanonicalEditPlan): OoxmlPackage {
  let pkg = input,
    allocation = 0;
  const prefix = `/oracle-target-${++epoch}`;
  const seed = (xml: string, partName = pkg.mainDocumentPart) => {
    const part = pkg.parts.get(partName)!;
    const parsed = readOoxmlPart(
      `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${xml}</w:body></w:document>`,
      { name: `${prefix}-${++allocation}.xml`, contentType: part.contentType }
    );
    if (!parsed.ok) throw new Error(parsed.reason);
    const body = bodyStoryRoot(parsed.part)!;
    return body.kind === 'textValue' ? [] : body.children;
  };
  const instances = new Map<string, string>(),
    levels = new Set<string>();
  const numId = (target: Extract<ListTarget, { kind: 'ordered' | 'bullet' }>): string => {
    if (target.instance.kind === 'existing') return target.instance.numId;
    let id = instances.get(target.instance.key);
    if (!id) {
      const made = ensureListDefinition(pkg, target.kind);
      if (!made) throw new Error('Cannot create oracle list');
      pkg = made.pkg;
      id = made.numId;
      instances.set(target.instance.key, id);
    }
    const levelKey = `${id}:${target.level}`;
    if (!levels.has(levelKey)) {
      const part = pkg.parts.get('/word/numbering.xml')!;
      const num = part.root.children.find(
        (node) =>
          node.kind !== 'textValue' &&
          node.localName === 'num' &&
          node.attributes.some((attr) => attr.localName === 'numId' && attr.value === id)
      )!;
      const start = target.instance.start ?? 1,
        format = target.kind === 'ordered' ? 'decimal' : 'bullet',
        text = target.kind === 'ordered' ? `%${target.level + 1}.` : '•',
        parsed = readOoxmlPart(
          `<w:numbering xmlns:w="${W}"><w:num w:numId="${id}"><w:lvlOverride w:ilvl="${target.level}"><w:startOverride w:val="${start}"/><w:lvl w:ilvl="${target.level}"><w:start w:val="${start}"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:pPr><w:ind w:left="${720 * (target.level + 1)}" w:hanging="360"/></w:pPr></w:lvl></w:lvlOverride></w:num></w:numbering>`,
          { name: `${prefix}-${++allocation}.xml`, contentType: part.contentType }
        );
      if (!parsed.ok) throw new Error('Invalid oracle numbering seed');
      const numbering = parsed.part.root.children[0]!;
      if (numbering.kind === 'textValue') throw new Error('Invalid oracle numbering');
      const inserted = insertChildren(
        part,
        num.id,
        num.kind === 'textValue' ? 0 : num.children.length,
        [numbering.children[0]!]
      );
      if (!inserted.ok) throw new Error('Invalid oracle numbering start');
      pkg = withPart(pkg, inserted.part);
      levels.add(levelKey);
    }
    return id;
  };
  const listXml = (target: ListTarget): string =>
    target.kind === 'ordered' || target.kind === 'bullet'
      ? `<w:numPr><w:ilvl w:val="${target.level}"/><w:numId w:val="${numId(target)}"/></w:numPr>`
      : target.kind === 'checklist'
        ? `<w:ind w:left="${720 * (target.level + 1)}" w:hanging="360"/>`
        : '';
  const marker = (checked: boolean) =>
    `<w:sdt><w:sdtPr><w:tag w:val="${checklistTag()}"/><w14:checkbox xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w14:checked w14:val="${checked ? 1 : 0}"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>${checked ? '☒' : '☐'}</w:t></w:r></w:sdtContent></w:sdt>`;
  const render = (node: TargetAstNode, partName: string, marks = ''): string => {
    const children = () =>
      ((node.children as TargetAstNode[]) ?? [])
        .map((child) => render(child, partName, marks))
        .join('');
    const run = (value: string) => `<w:r>${marks ? `<w:rPr>${marks}</w:rPr>` : ''}${value}</w:r>`;
    if (node.type === 'text')
      return run(`<w:t xml:space="preserve">${escapeXml(String(node.value))}</w:t>`);
    if (node.type === 'break') return run('<w:br/>');
    if (['strong', 'emphasis', 'delete'].includes(node.type))
      return (node.children as TargetAstNode[])
        .map((child) =>
          render(
            child,
            partName,
            marks +
              (node.type === 'strong'
                ? '<w:b/>'
                : node.type === 'emphasis'
                  ? '<w:i/>'
                  : '<w:strike/>')
          )
        )
        .join('');
    if (node.type === 'inlineCode')
      return `<w:r><w:rPr>${marks}<w:rStyle w:val="Code"/></w:rPr><w:t>${escapeXml(String(node.value))}</w:t></w:r>`;
    if (node.type === 'link') {
      const link = ensureHyperlinkRelationship(pkg, String(node.url), partName);
      if (!link) throw new Error('Unsafe oracle link');
      pkg = link.pkg;
      return `<w:hyperlink r:id="${escapeXml(link.relationshipId)}">${children()}</w:hyperlink>`;
    }
    throw new Error(`Unsupported oracle inline: ${node.type}`);
  };
  for (const edit of plan.operations) {
    if (edit.kind === 'insert-column') {
      pkg = materializeInsertedColumn(pkg, { ...plan, operations: [edit] }).pkg;
    } else if (edit.kind === 'insert-row') {
      pkg = materializeInsertedRow(pkg, { ...plan, operations: [edit] }).pkg;
    } else if (edit.kind === 'table') {
      const part = pkg.parts.get(edit.partName)!,
        result = applyTreeOp(part, edit.operation);
      if (!result.ok) throw new Error(result.reason);
      pkg = withPart(pkg, result.part);
    } else if (edit.kind === 'siblings') {
      let part = pkg.parts.get(edit.partName)!;
      const parent = findNode(part, edit.parentId)!;
      if (parent.kind === 'textValue') throw new Error('Invalid oracle parent');
      let at =
        edit.afterId === null
          ? parent.children.length
          : parent.children.findIndex((node) => node.id === edit.afterId);
      if (edit.removedIds.length)
        at = parent.children.findIndex((node) => node.id === edit.removedIds[0]);
      if (at < 0) throw new Error('Missing oracle anchor');
      const lists = new Map(
        (edit.lists ?? []).map((binding) => {
          if (binding.path.length !== 1) throw new Error('Oracle requires root paragraph paths');
          return [binding.path[0]!, binding.target] as const;
        })
      );
      const xml = edit.inserted
        .map((ast, index) => {
          if (!['paragraph', 'heading'].includes(ast.type))
            throw new Error('Unsupported oracle block');
          const target = lists.get(index) ?? { kind: 'none' };
          const style =
            ast.type === 'heading'
              ? `<w:pStyle w:val="Heading${ast.depth}"/><w:outlineLvl w:val="${Number(ast.depth) - 1}"/>`
              : '';
          return `<w:p><w:pPr>${style}${listXml(target)}</w:pPr>${target.kind === 'checklist' ? marker(target.checked) : ''}${(ast.children as TargetAstNode[]).map((child) => render(child, edit.partName)).join('')}</w:p>`;
        })
        .join('');
      if (xml) {
        const added = insertChildren(part, edit.parentId, at, seed(xml, edit.partName));
        if (!added.ok) throw new Error('Oracle insertion failed');
        part = added.part;
      }
      for (const id of edit.removedIds) {
        const removed = removeNode(part, id);
        if (!removed.ok) throw new Error('Oracle removal failed');
        part = removed.part;
      }
      pkg = withPart(pkg, part);
    } else if (edit.kind === 'list') {
      let part = pkg.parts.get(edit.partName)!;
      const apply = (op: Parameters<typeof applyTreeOp>[1]) => {
        const result = applyTreeOp(part, op);
        if (!result.ok) throw new Error(result.reason);
        part = result.part;
      };
      const target = edit.after;
      if (edit.before.checkbox && target.kind !== 'checklist')
        apply({
          op: 'removeContentControl',
          controlId: edit.before.checkbox.controlId,
          keepContent: false,
        });
      if (target.kind === 'checklist') {
        if (edit.before.checkbox)
          apply({
            op: 'setContentControlValue',
            controlId: edit.before.checkbox.controlId,
            value: { kind: 'checkbox', checked: target.checked },
          });
        else {
          const paragraph = findNode(part, edit.paragraphId)!;
          if (paragraph.kind !== 'paragraph') throw new Error('Missing oracle paragraph');
          const wrapper = seed(`<w:p>${marker(target.checked)}</w:p>`)[0]!;
          const control =
            wrapper.kind === 'paragraph' &&
            wrapper.children.find((node) => node.kind === 'contentControl');
          if (!control || !contentControlPropertiesOf(control).checkbox)
            throw new Error('Invalid oracle checkbox');
          const inserted = insertChildren(
            part,
            paragraph.id,
            paragraph.children[0]?.localName === 'pPr' ? 1 : 0,
            [control]
          );
          if (!inserted.ok) throw new Error('Oracle checkbox insertion failed');
          part = inserted.part;
        }
        apply({ op: 'setListNumbering', paragraphId: edit.paragraphId, numId: '0' });
      } else
        apply({
          op: 'setListNumbering',
          paragraphId: edit.paragraphId,
          numId: target.kind === 'none' ? '0' : numId(target),
          ...(target.kind === 'none' ? {} : { level: target.level }),
        });
      if (target.kind === 'checklist' || edit.before.checkbox) {
        const properties = directParagraphProperties(part, edit.paragraphId),
          attributes = {
            ...properties.find((property) => property.localName === 'ind')?.attributes,
          };
        for (const name of ['left', 'start', 'hanging']) delete attributes[name];
        if (target.kind === 'checklist') {
          for (const name of [
            'firstLine',
            'firstLineChars',
            'hangingChars',
            'leftChars',
            'startChars',
          ])
            delete attributes[name];
          attributes.left = String(720 * (target.level + 1));
          attributes.hanging = '360';
        }
        apply({
          op: 'setParagraphProperties',
          paragraphId: edit.paragraphId,
          properties: [
            ...properties.filter((property) => property.localName !== 'ind'),
            ...(Object.keys(attributes).length ? [{ localName: 'ind', attributes }] : []),
          ],
        });
      }
      pkg = withPart(pkg, part);
    } else pkg = applyTargetRetained(pkg, { ...plan, operations: [edit] });
  }
  return pkg;
}
