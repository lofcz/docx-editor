import { useLayoutEffect, useRef } from 'react';
import {
  revealRibbonTab,
  ribbonKeyTab,
  ribbonTableCommands,
  ribbonTabs,
  type RibbonTabControls,
} from '@docx-editor.dev/core/editor';
import { useEditorCommand } from '../useEditorCommand';
import { useToolbarLabel } from './toolbar-context';

export function RibbonTabs({
  tabs,
  active,
  panelId,
  collapsed,
  rtl,
  label,
  select,
  toggle,
}: RibbonTabControls) {
  const selected = useRef<HTMLButtonElement>(null);
  useLayoutEffect(() => {
    if (selected.current) revealRibbonTab(selected.current);
  }, [active]);
  const disclosure = label(collapsed ? 'ribbon.expand' : 'ribbon.collapse');
  return (
    <div className="docx-ribbon__header">
      <div className="docx-ribbon__tabs" role="tablist" aria-label={label('ribbon.label')}>
        {tabs.map((tab) => (
          <button
            key={tab}
            ref={tab === active ? selected : undefined}
            type="button"
            role="tab"
            id={panelId + '-' + tab}
            aria-controls={panelId}
            aria-selected={tab === active}
            tabIndex={tab === active ? 0 : -1}
            data-ribbon-tab={tab}
            onClick={() => select(tab)}
            onKeyDown={(event) => {
              const next = ribbonKeyTab(event.key, tabs, tab, rtl);
              if (!next) return;
              event.preventDefault();
              const target = event.currentTarget.parentElement?.querySelector<HTMLButtonElement>(
                `[data-ribbon-tab="${next}"]`
              );
              target?.focus({ preventScroll: true });
              select(next);
            }}
          >
            {label(ribbonTabs[tab])}
          </button>
        ))}
      </div>
      <button
        className="docx-ribbon__toggle"
        type="button"
        aria-expanded={!collapsed}
        aria-controls={panelId}
        aria-label={disclosure}
        title={disclosure}
        onClick={toggle}
      >
        <span aria-hidden="true" />
      </button>
    </div>
  );
}

function commandPart(key: keyof typeof ribbonTableCommands) {
  return function RibbonCommand() {
    const command = useEditorCommand(ribbonTableCommands[key]);
    const label = useToolbarLabel()(key);
    return (
      <button
        type="button"
        className="docx-ribbon__command"
        data-ribbon-command={key}
        disabled={!command.isEnabled}
        title={command.disabledReason ?? label}
        onClick={() => command.execute()}
      >
        {label}
      </button>
    );
  };
}
type CommandKey = keyof typeof ribbonTableCommands;
export const ribbonCommandParts = Object.fromEntries(
  (Object.keys(ribbonTableCommands) as CommandKey[]).map((key) => [key, commandPart(key)])
) as Record<CommandKey, ReturnType<typeof commandPart>>;
