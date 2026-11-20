// The ribbon arranges existing slots. Its table actions use the editor's normal command probes.
import type { EditorCommand, EditorSnapshot } from '../contracts/editor.ts';
import {
  chromeSlotId,
  formattingBarChromeGroups,
  type ChromeSlotId,
  type FormattingBarOptions,
} from './chrome-controls.ts';

/** @internal */
export const ribbonTableCommands = {
  'ribbon.blueGrid': { type: 'setTableProperties', styleId: 'BandedGridBlue' },
  'ribbon.greenGrid': { type: 'setTableProperties', styleId: 'BandedGridGreen' },
  'table.insertRowAbove': { type: 'insertRow', where: 'above' },
  'table.insertRowBelow': { type: 'insertRow', where: 'below' },
  'table.insertColumnLeft': { type: 'insertColumn', where: 'left' },
  'table.insertColumnRight': { type: 'insertColumn', where: 'right' },
  'table.deleteRow': { type: 'deleteRow' },
  'table.deleteColumn': { type: 'deleteColumn' },
  'tableAdvanced.autoFit': { type: 'setTableProperties', autoFit: 'contents' },
  'ribbon.fitWindow': { type: 'setTableProperties', autoFit: 'window' },
  'ribbon.fixedWidth': { type: 'setTableProperties', autoFit: 'fixed' },
} as const satisfies Record<string, EditorCommand>;

/** @internal */
export type RibbonSlot = ChromeSlotId | 'alignment' | keyof typeof ribbonTableCommands;

/** @internal */
export const ribbonTabs = {
  home: 'ribbon.home',
  insert: 'toolbar.insert',
  layout: 'ribbon.layout',
  review: 'toolbar.review',
  view: 'ribbon.view',
  picture: 'ribbon.pictureFormat',
  tableDesign: 'ribbon.tableDesign',
  tableLayout: 'ribbon.tableLayout',
} as const;

/** @internal */
export type RibbonTab = keyof typeof ribbonTabs;

/** @internal */
export interface RibbonGroup {
  id: string;
  tab: RibbonTab;
  labelKey: string;
  slots: readonly RibbonSlot[];
}

const group = (
  tab: RibbonTab,
  id: string,
  labelKey: string,
  ...slots: RibbonSlot[]
): RibbonGroup => ({ id, tab, labelKey, slots });

const additionalGroups = [
  group('insert', 'insertTables', 'ribbon.tables', 'table.insert'),
  group(
    'insert',
    'illustrations',
    'ribbon.illustrations',
    'image.insert',
    'insert.chart',
    'insert.textBox'
  ),
  group(
    'insert',
    'references',
    'ribbon.references',
    'insert.footnote',
    'insert.endnote',
    'insert.bookmark'
  ),
  group(
    'layout',
    'breaks',
    'ribbon.breaks',
    'insert.pageBreak',
    'insert.sectionBreakNextPage',
    'insert.sectionBreakContinuous'
  ),
  group(
    'review',
    'changes',
    'ribbon.changes',
    'review.previousChange',
    'review.nextChange',
    'review.revisionMarkup'
  ),
  group('view', 'show', 'ribbon.show', 'review.paragraphMarks', 'review.bookmarks'),
];
const tableGroups = [
  group('tableDesign', 'tableStyles', 'ribbon.tableStyles', 'ribbon.blueGrid', 'ribbon.greenGrid'),
  group(
    'tableDesign',
    'tableBorders',
    'ribbon.bordersAndShading',
    'table.borderTarget',
    'table.borderColor',
    'table.borderStyle',
    'table.borderWidth',
    'table.cellFill'
  ),
  group(
    'tableLayout',
    'tableRows',
    'ribbon.rows',
    'table.insertRowAbove',
    'table.insertRowBelow',
    'table.deleteRow'
  ),
  group(
    'tableLayout',
    'tableColumns',
    'ribbon.columns',
    'table.insertColumnLeft',
    'table.insertColumnRight',
    'table.deleteColumn'
  ),
  group('tableLayout', 'tableMerge', 'ribbon.merge', 'table.mergeCells', 'table.splitCell'),
  group(
    'tableLayout',
    'tableSize',
    'ribbon.cellSize',
    'tableAdvanced.autoFit',
    'ribbon.fitWindow',
    'ribbon.fixedWidth'
  ),
];

/** @internal */
export function ribbonGroups(
  image: EditorSnapshot['image'],
  options: FormattingBarOptions,
  table: boolean
): readonly RibbonGroup[] {
  const groups = formattingBarChromeGroups(image, options).map((item): RibbonGroup => {
    const object = ['image', 'shape', 'chart'].includes(item.id);
    const tab = object
      ? 'picture'
      : item.id === 'review'
        ? 'review'
        : item.id === 'zoom'
          ? 'view'
          : 'home';
    return {
      id: item.id,
      labelKey: item.labelKey,
      tab,
      slots:
        item.id === 'alignment'
          ? ['alignment']
          : item.controls
              .map((control) => chromeSlotId(item, control))
              .filter((slot) => slot !== 'image.insert'),
    };
  });
  return [...groups, ...additionalGroups, ...(table ? tableGroups : [])];
}

/** @internal */
export function ribbonKeyTab(
  key: string,
  tabs: readonly RibbonTab[],
  active: RibbonTab,
  rtl: boolean
): RibbonTab | undefined {
  if (key === 'Home') return tabs[0];
  if (key === 'End') return tabs.at(-1);
  if (key !== 'ArrowLeft' && key !== 'ArrowRight') return undefined;
  const step = (key === 'ArrowRight') !== rtl ? 1 : -1;
  return tabs[(tabs.indexOf(active) + step + tabs.length) % tabs.length];
}

/** Move only the tab strip. Focusing a tab must not move the document viewport. @internal */
export function revealRibbonTab(button: HTMLElement): void {
  const strip = button.parentElement;
  if (!strip) return;
  const outer = strip.getBoundingClientRect();
  const tab = button.getBoundingClientRect();
  if (tab.left < outer.left) strip.scrollLeft += tab.left - outer.left;
  if (tab.right > outer.right) strip.scrollLeft += tab.right - outer.right;
}

/** @internal */
export function isRibbonTableCommand(slot: RibbonSlot): slot is keyof typeof ribbonTableCommands {
  return Object.hasOwn(ribbonTableCommands, slot);
}

/** @internal */
export function ribbonSelection(
  groups: readonly { tab: RibbonTab }[],
  preferred: RibbonTab | undefined,
  context: RibbonTab
) {
  const tabs = (Object.keys(ribbonTabs) as RibbonTab[]).filter((tab) =>
    groups.some((g) => g.tab === tab)
  );
  if (!tabs.length) tabs.push('home');
  const active =
    preferred && tabs.includes(preferred) ? preferred : tabs.includes(context) ? context : tabs[0]!;
  return { tabs, active };
}

/** @internal */
export interface RibbonTabControls {
  tabs: readonly RibbonTab[];
  active: RibbonTab;
  panelId: string;
  collapsed: boolean;
  rtl: boolean;
  label: (key: string) => string;
  select: (tab: RibbonTab) => void;
  toggle: () => void;
}

/** Combine adjacent registry groups under their Word ribbon caption, after host placement and overflow. @internal */
export function ribbonBands<T>(groups: readonly { id: string; node: T }[]) {
  const bands: { id: string; caption?: string; nodes: T[] }[] = [];
  for (const { id, node } of groups) {
    const caption = ['font', 'text', 'script'].includes(id)
      ? 'formattingBar.groups.font'
      : ['alignment', 'direction', 'list'].includes(id)
        ? 'dialogs.paragraph.title'
        : undefined;
    const last = bands.at(-1);
    if (caption && last?.caption === caption) last.nodes.push(node);
    else bands.push({ id, caption, nodes: [node] });
  }
  return bands;
}
