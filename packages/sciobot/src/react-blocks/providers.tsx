import type { ReactNode } from 'react';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ReactBlockViewProps } from './host';

export interface ReactBlockProviderScope {
  readonly editor: DocxEditorInstance;
  readonly block: ReactBlockViewProps['block'];
  readonly fragment: ReactBlockViewProps['fragment'];
}

/** Return provider elements. Keep their component types stable; call hooks inside components. */
export type ReactBlockProviders = (
  children: ReactNode,
  scope: ReactBlockProviderScope
) => ReactNode;

export interface ReactBlockHostOptions {
  readonly providers?: ReactBlockProviders;
  /** Widget height in CSS pixels for `reserveHeight` roots. Defaults to the client rect. */
  readonly measure?: import('./measurement').MeasureBlockElement;
}

/** The stable root boundary retains widget identity when provider callbacks or values change. */
export function BlockRootProviders({
  children,
  render,
  scope,
}: {
  readonly children: ReactNode;
  readonly render: ReactBlockProviders | undefined;
  readonly scope: ReactBlockProviderScope;
}) {
  return render ? render(children, scope) : children;
}
