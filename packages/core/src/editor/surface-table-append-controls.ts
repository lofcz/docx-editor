// Table append controls share native table plans and layout geometry with resize furniture.
import type { EditorCommand } from '../contracts/editor.ts';
import { directParagraphsInCells, tableContextAt } from '../layout/semantic-cell-selection.ts';
import type { TableFragmentRecord } from '../layout/semantic-records.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import { planTableCommand, type TableCommandPlan } from './table-command-plan.ts';
import { findTableOnPage, partOfTableIn } from './surface-table-story.ts';
import { sameStoryPart } from './surface-scope.ts';
import type {
  SurfaceTableInteractionHost,
  SurfaceTableInteractionInput,
} from './surface-table-interaction.ts';

type AppendKind = 'row' | 'column' | 'both';
interface Target {
  readonly table: TableFragmentRecord;
  readonly pageIndex: number;
  readonly origin: { readonly x: number; readonly y: number };
}

/** The last painted fragment owns append controls; repeated headers never become targets. */
function targetOf(input: SurfaceTableInteractionInput, tableId: string): Target | null {
  for (let pageIndex = input.layout.pages.length - 1; pageIndex >= 0; pageIndex--) {
    const found = findTableOnPage(input.layout, tableId, pageIndex);
    if (found && found.table.rows.some((row) => !row.isHeaderRepeat))
      return { ...found, pageIndex };
  }
  return null;
}

export function createTableAppendControls(host: SurfaceTableInteractionHost) {
  const layer = host.furnitureLayer.ownerDocument.createElement('div');
  layer.className = 'docx-table-append-controls';
  layer.setAttribute('contenteditable', 'false');
  layer.setAttribute('data-docx-marker', '');
  Object.assign(layer.style, { position: 'absolute', left: '0', top: '0', pointerEvents: 'none' });
  host.furnitureLayer.after(layer);
  let target: Target | null = null;
  let sourceRevision = -1;
  let painted: { layout: SurfaceTableInteractionInput['layout']; key: string } | null = null;
  const buttons = new Map<AppendKind, HTMLButtonElement>();

  const label = (kind: AppendKind) =>
    kind === 'row'
      ? host.label('table.insertRowBelow')
      : kind === 'column'
        ? host.label('table.insertColumnRight')
        : `${host.label('table.insertRowBelow')} + ${host.label('table.insertColumnRight')}`;

  function plan(kind: AppendKind, input: SurfaceTableInteractionInput): TableCommandPlan {
    const refusal = (reason: string): TableCommandPlan => ({
      ok: false,
      code: 'unsupported',
      reason,
    });
    if (!target || sourceRevision !== input.storeRevision)
      return refusal('the table target is no longer valid');
    const tableId = target.table.tableId;
    const owned = host.controlRefusal?.(tableId);
    if (owned) return refusal(owned);
    const part = partOfTableIn(host.session(), tableId);
    const topology = readEditableTableTopology(part.root, tableId);
    if (!topology.ok) return refusal('the table target is no longer valid');
    const row = topology.topology.rows.at(-1)?.row;
    if (!row) return refusal('the table has no rows');
    // An RTL table's visual right edge names its first logical grid column.
    const rtl = target.table.rows.some((item) =>
      item.cells.some(
        (cell) => cell.logicalGridColumn !== undefined && cell.logicalGridColumn !== cell.gridColumn
      )
    );
    const gridColumn = rtl
      ? topology.topology.gridColumns[0]
      : topology.topology.gridColumns.at(-1);
    const commands: EditorCommand[] = [];
    let selection = input.selection;
    let cellSelection = input.cellSelection;
    if (kind !== 'column')
      commands.push({
        type: 'insertRow',
        where: 'below',
        target: { sourceRevision, tableId, rowId: row.id, isHeaderRepeat: false },
      });
    if (kind !== 'row') {
      if (gridColumn) {
        commands.push({
          type: 'insertColumn',
          where: 'right',
          target: { sourceRevision, tableId, gridColumnId: gridColumn.id, isHeaderRepeat: false },
        });
      } else {
        // Word permits tables without tblGrid. Use the native planner's cell-reference path.
        // This planning position never changes the editor's actual caret or cell selection.
        const cells = topology.topology.rows[0]?.cells;
        const cell = rtl ? cells?.[0] : cells?.at(-1);
        const paragraphId = cell && directParagraphsInCells(input.layout, [cell.id])[0];
        if (!paragraphId) return refusal('the table edge has no editable cell');
        const position = { paragraphId, offset: 0 };
        selection = { anchor: position, head: position };
        cellSelection = null;
        commands.push({ type: 'insertColumn', where: 'right' });
      }
    }
    const plans = commands.map((command) =>
      planTableCommand({
        command,
        part,
        layout: input.layout,
        storeRevision: input.storeRevision,
        selection,
        cellSelection,
        themeColors: input.themeColors,
        editable: host.session().editable && input.editingMode !== 'view',
        suggestingAuthor: input.suggestingAuthor,
        viewing: input.editingMode === 'view',
      })
    );
    for (const next of plans) if (!next.ok) return next;
    // These canonical operations resolve their own row cells when the transaction applies.
    // The column operation therefore includes the row that the preceding operation inserts.
    const ops = plans.flatMap((next) => (next.ok ? next.ops : []));
    const reason = host.planRefusal?.(ops);
    if (reason) return refusal(reason);
    return {
      ok: true,
      ops,
      selection: { kind: 'adoptCommittedCaret' },
    };
  }

  function clear() {
    const focused = layer.contains(layer.ownerDocument.activeElement);
    layer.replaceChildren();
    buttons.clear();
    target = null;
    painted = null;
    if (focused) host.pagesLayer.focus({ preventScroll: true });
  }

  function button(kind: AppendKind): HTMLButtonElement {
    const old = buttons.get(kind);
    if (old) return old;
    const next = layer.ownerDocument.createElement('button');
    next.type = 'button';
    next.className = `docx-table-append-${kind}`;
    next.dataset.tableAppend = kind;
    const icon = layer.ownerDocument.createElement('span');
    icon.setAttribute('aria-hidden', 'true');
    icon.textContent = '+';
    next.append(icon);
    Object.assign(next.style, { position: 'absolute', pointerEvents: 'auto' });
    next.addEventListener('mousedown', (event) => event.preventDefault());
    next.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      const keyboard = layer.ownerDocument.activeElement === next;
      const result = host.applyTableCommandPlan(plan(kind, host.read()));
      if (!result.ok) {
        next.title = result.reason;
        return;
      }
      if (!keyboard) host.pagesLayer.focus({ preventScroll: true });
    });
    next.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        host.pagesLayer.focus({ preventScroll: true });
      }
    });
    buttons.set(kind, next);
    layer.append(next);
    return next;
  }

  function update(input: SurfaceTableInteractionInput, hoveredTableId?: string) {
    const tableId =
      hoveredTableId ??
      input.cellSelection?.tableId ??
      tableContextAt(input.layout, input.selection.head.paragraphId)?.tableId;
    if (
      !tableId ||
      input.editingMode === 'view' ||
      host.controlRefusal?.(tableId) != null ||
      !sameStoryPart(host.session(), tableId, input.selection.head.paragraphId)
    ) {
      clear();
      return;
    }
    const key = `${tableId}:${input.storeRevision}:${input.editingMode}:${input.suggestingAuthor ?? ''}:${host.scale()}`;
    if (painted?.layout === input.layout && painted.key === key) return;
    painted = { layout: input.layout, key };
    const found = targetOf(input, tableId);
    if (!found) {
      clear();
      return;
    }
    target = found;
    sourceRevision = input.storeRevision;
    const scale = host.scale();
    const left =
      (target.origin.x + target.table.box.x + host.pageOffsetX(target.pageIndex)) * scale;
    const top = (target.origin.y + target.table.box.y) * scale;
    const width = target.table.box.width * scale;
    const height = target.table.box.height * scale;
    layer.style.width = host.pagesLayer.style.width;
    layer.style.height = host.pagesLayer.style.height;
    for (const kind of ['row', 'column', 'both'] as const) {
      const control = button(kind);
      const verdict = plan(kind, input);
      control.disabled = !verdict.ok;
      control.setAttribute('aria-label', label(kind));
      control.title = verdict.ok ? label(kind) : verdict.reason;
      Object.assign(control.style, {
        left: `${kind === 'row' ? left - 26 : left + width + 2}px`,
        top: `${kind === 'column' ? top + Math.max(0, height / 2 - 12) : top + height - 24}px`,
        width: '24px',
        height: '24px',
      });
    }
  }
  return {
    update,
    refreshLabels: () => {
      for (const [kind, control] of buttons) {
        control.setAttribute('aria-label', label(kind));
        if (!control.disabled) control.title = label(kind);
      }
    },
    contains: (node: Node | null) => node !== null && layer.contains(node),
    destroy: () => layer.remove(),
  };
}
