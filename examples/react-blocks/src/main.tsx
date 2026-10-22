import { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor, useDocxEditor, useEditorState } from '@docx-editor.dev/react';
import { loadDefaultFonts } from '@docx-editor.dev/fonts';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { fixture, readBlocks } from './document';
import { REVIEW, createBlockBridge } from './bridge';
import { mountReactBlocks } from './react-block-layer';
import { PhaseBlock, ActivityBlock, QuizBlock } from './blocks';
import '@docx-editor.dev/core/styles/editor.css';
import './style.css';

const MODULES = [REVIEW];
const REGISTRY = { phase: PhaseBlock, activity: ActivityBlock, quiz: QuizBlock };
const INITIAL = fixture();
const FONTS = await loadDefaultFonts({ families: ['Calibri'] });
if (FONTS.failures.length) throw new Error('Packaged document fonts failed to load');

function Proof({ reload }: { reload: (bytes: ArrayBuffer) => void }) {
  const editor = useDocxEditor() as DocxEditorInstance | null;
  const snapshot = useEditorState((s) => s);
  const [bridge, setBridge] = useState<ReturnType<typeof createBlockBridge> | null>(null);
  const [error, setError] = useState('');
  useEffect(() => {
    if (!editor?.surface) return;
    const adapter = createBlockBridge(editor);
    const stop = mountReactBlocks(editor, REGISTRY, adapter);
    setBridge(adapter);
    const api = {
      editor,
      blocks: () => readBlocks(editor.surface!.session.part()),
      save: () => editor.save(),
      load: reload,
      propose: () => adapter.propose(),
    };
    Object.assign(window, { proof: api });
    return () => {
      stop();
      adapter.destroy();
    };
  }, [editor, reload]);
  const run = (action: () => unknown) => {
    try {
      action();
      setError('');
    } catch (error) {
      setError(String(error));
    }
  };
  const changes = editor?.getReviewItems().filter((item) => item.kind === 'revision') ?? [];
  return (
    <>
      <div className="proof-toolbar">
        <button disabled={!bridge} onClick={() => run(() => bridge!.propose())}>
          Propose AI text edit
        </button>
        <button
          disabled={!changes.length}
          onClick={() =>
            run(() => {
              for (const item of editor!.getReviewItems()) editor!.acceptReviewItem(item.key);
            })
          }
        >
          Accept changes
        </button>
        <button
          disabled={!changes.length}
          onClick={() =>
            run(() => {
              for (const item of editor!.getReviewItems()) editor!.rejectReviewItem(item.key);
            })
          }
        >
          Reject changes
        </button>
        <button onClick={() => run(() => editor!.exec({ type: 'undo' }))}>Undo</button>
        <button onClick={() => run(() => editor!.exec({ type: 'redo' }))}>Redo</button>
        <button disabled={!bridge} onClick={() => run(() => bridge!.pushBlocks())}>
          Push blocks to next page
        </button>
        <button
          onClick={() =>
            void editor
              ?.save()
              .then(reload)
              .catch((e) => setError(String(e)))
          }
        >
          Save and reopen DOCX
        </button>
        <span data-testid="review-count">{changes.length} pending changes</span>
        <span data-testid="page-count">{snapshot?.page.total ?? 0} pages</span>
        {error ? <strong role="alert">{error}</strong> : null}
      </div>
      <DocxEditor.Viewport>
        <DocxEditor.Content />
      </DocxEditor.Viewport>
    </>
  );
}
function App() {
  const [document, setDocument] = useState<Uint8Array | ArrayBuffer>(INITIAL);
  return (
    <main>
      <h1>Arbitrary React blocks proof</h1>
      <p>
        React controls, nested containers, editable Word paragraphs, native tables, tracked text
        changes, and DOCX persistence.
      </p>
      <DocxEditor.Root document={document} modules={MODULES} fonts={FONTS} author="ScioBot">
        <Proof reload={setDocument} />
      </DocxEditor.Root>
    </main>
  );
}
createRoot(window.document.getElementById('app')!).render(<App />);
