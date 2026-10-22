import { createElement, useLayoutEffect, useRef, type ReactNode } from 'react';
import type { DocxEditorInstance, PaginatedSurface } from '@docx-editor.dev/core/editor';
import { mountReactBlockHost, type ReactBlockView } from './host';
import { InlineMeasurement, type MeasureInlineElement } from './inline-measurement';
import type { ReactBlockProviders } from './providers';

export interface ReactInlineHostOptions {
  readonly providers?: ReactBlockProviders;
  readonly measure?: MeasureInlineElement;
}
export interface ReactInlineHost {
  register(type: string, view: ReactBlockView): () => void;
  setProviders(providers: ReactBlockProviders | undefined): void;
  refresh(): void;
  /** Dispose before destroying the editor. React cleanup completes in the next microtask. */
  dispose(): void;
}
const owners = new WeakSet<PaginatedSurface>();

/** Replace registered inline fallbacks with React views inside native atomic layout boxes. */
export function mountReactInlineHost(
  editor: DocxEditorInstance,
  initial: ReadonlyMap<string, ReactBlockView> = new Map(),
  options: ReactInlineHostOptions = {}
): ReactInlineHost {
  const surface = editor.surface;
  if (!surface) throw new Error('Attach the editor before mounting React inline widgets.');
  if (owners.has(surface)) throw new Error('The surface already has a React inline host.');
  const measurement = new InlineMeasurement(surface, options.measure);
  const wrap = (View: ReactBlockView): ReactBlockView =>
    function InlineView(props) {
      return createElement(MeasuredInlineRoot, {
        measurement,
        controlId: props.block.controlId,
        zoom: props.fragment.scale * 0.75,
        children: createElement(View, props),
      });
    };
  const host = mountReactBlockHost(
    editor,
    new Map([...initial].map(([type, view]) => [type, wrap(view)])),
    { providers: options.providers, inlineOnly: true }
  );
  owners.add(surface);
  const stop = surface.observePaint({
    beforePaint: () => measurement.beforePaint(),
    afterPaint: (frame) => measurement.place(frame),
    destroyed: dispose,
  });
  let disposed = false;
  function dispose(): void {
    if (disposed) return;
    disposed = true;
    stop();
    host.dispose();
    measurement.dispose();
    owners.delete(surface!);
  }
  return {
    register(type, view) {
      const release = host.register(type, wrap(view));
      measurement.place();
      return () => {
        release();
        measurement.place();
      };
    },
    setProviders(providers) {
      host.setProviders(providers);
      measurement.place();
    },
    refresh() {
      host.refresh();
      measurement.place();
    },
    dispose,
  };
}

function MeasuredInlineRoot({
  measurement,
  controlId,
  zoom,
  children,
}: {
  readonly measurement: InlineMeasurement;
  readonly controlId: string;
  readonly zoom: number;
  readonly children: ReactNode;
}) {
  const ref = useRef<HTMLSpanElement>(null),
    baseline = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    if (ref.current && baseline.current)
      measurement.track(controlId, ref.current, baseline.current, zoom);
  });
  useLayoutEffect(() => {
    const element = ref.current;
    return () => {
      if (element) measurement.release(controlId, element);
    };
  }, [measurement, controlId]);
  return createElement(
    'span',
    {
      ref,
      'data-sciobot-inline-measure': controlId,
      style: {
        display: 'inline-flex',
        alignItems: 'baseline',
        whiteSpace: 'nowrap',
        transform: `scale(${zoom})`,
        transformOrigin: '0 0',
      },
    },
    createElement('span', { style: { display: 'inline-block' } }, children),
    createElement('span', {
      ref: baseline,
      'data-sciobot-inline-baseline': '',
      style: { display: 'inline-block', width: 0, height: 0, padding: 0, border: 0 },
    })
  );
}
