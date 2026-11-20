import { defineComponent, ref, watchPostEffect, type PropType } from 'vue';
import {
  revealRibbonTab,
  ribbonKeyTab,
  ribbonTableCommands,
  ribbonTabs,
  type RibbonTab,
  type RibbonTabControls,
} from '@docx-editor.dev/core/editor';
import { useEditorCommand } from '../useEditorCommand';
import { useToolbarLabel } from './toolbar-context';

export const RibbonTabs = defineComponent({
  name: 'RibbonTabs',
  props: {
    tabs: { type: Array as PropType<readonly RibbonTab[]>, required: true },
    active: { type: String as PropType<RibbonTab>, required: true },
    panelId: { type: String, required: true },
    collapsed: Boolean,
    rtl: Boolean,
    label: { type: Function as PropType<RibbonTabControls['label']>, required: true },
    select: { type: Function as PropType<RibbonTabControls['select']>, required: true },
    toggle: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const strip = ref<HTMLElement | null>(null);
    watchPostEffect(() => {
      const selected = strip.value?.querySelector<HTMLElement>(
        `[data-ribbon-tab="${props.active}"]`
      );
      if (selected) revealRibbonTab(selected);
    });
    return () => {
      const disclosure = props.label(props.collapsed ? 'ribbon.expand' : 'ribbon.collapse');
      return (
        <div class="docx-ribbon__header">
          <div
            class="docx-ribbon__tabs"
            ref={strip}
            role="tablist"
            aria-label={props.label('ribbon.label')}
          >
            {props.tabs.map((tab) => (
              <button
                key={tab}
                type="button"
                role="tab"
                id={props.panelId + '-' + tab}
                aria-controls={props.panelId}
                aria-selected={tab === props.active}
                tabindex={tab === props.active ? 0 : -1}
                data-ribbon-tab={tab}
                onClick={() => props.select(tab)}
                onKeydown={(event: KeyboardEvent) => {
                  const next = ribbonKeyTab(event.key, props.tabs, tab, props.rtl);
                  if (!next) return;
                  event.preventDefault();
                  strip.value
                    ?.querySelector<HTMLButtonElement>(`[data-ribbon-tab="${next}"]`)
                    ?.focus({ preventScroll: true });
                  props.select(next);
                }}
              >
                {props.label(ribbonTabs[tab])}
              </button>
            ))}
          </div>
          <button
            class="docx-ribbon__toggle"
            type="button"
            aria-expanded={!props.collapsed}
            aria-controls={props.panelId}
            aria-label={disclosure}
            title={disclosure}
            onClick={props.toggle}
          >
            <span aria-hidden="true" />
          </button>
        </div>
      );
    };
  },
});

function commandPart(key: keyof typeof ribbonTableCommands) {
  return defineComponent({
    name: 'RibbonCommand',
    setup() {
      const command = useEditorCommand(ribbonTableCommands[key]);
      const label = useToolbarLabel();
      return () => (
        <button
          type="button"
          class="docx-ribbon__command"
          data-ribbon-command={key}
          disabled={!command.isEnabled.value}
          title={command.disabledReason.value ?? label(key)}
          onClick={() => command.execute()}
        >
          {label(key)}
        </button>
      );
    },
  });
}
type CommandKey = keyof typeof ribbonTableCommands;
export const ribbonCommandParts = Object.fromEntries(
  (Object.keys(ribbonTableCommands) as CommandKey[]).map((key) => [key, commandPart(key)])
) as Record<CommandKey, ReturnType<typeof commandPart>>;
