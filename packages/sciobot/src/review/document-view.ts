import type { MarkdownLayoutOptions, MarkdownImageAsset } from '@docx-editor.dev/docx-to-markdown';
import { nativeImageAssets } from '../projection/inline-image-assets';
import {
  openDocumentForExport,
  type ExportSemanticLayout,
  type OpenDocumentForExportOptions,
} from '@docx-editor.dev/core/export';
import {
  TreePackageStore,
  headlessViewOfStore,
  type HeadlessDocumentView,
} from '@docx-editor.dev/core/store';
import {
  sameRevision,
  type DocumentView,
  type NativeEditFailure,
  type RevisionToken,
} from '../contracts/document';
import { projectMdx, type MdxProjection } from '../projection/mdx';
import { projectPendingProperties, type PendingPropertyProposal } from './property-projection';
import { readApplicationProposals, projectApplicationPresence } from './application-proposals';
import { withoutTaskCompletionAppearance } from '../native/task-completion-style';

export type ReviewViewOptions = Pick<
  OpenDocumentForExportOptions,
  'measurer' | 'producer' | 'resourceTimeoutMs' | 'signal'
> & {
  /** Optional property/atomic snapshot. Reads also load stored proposals without resolving them. */
  readonly pendingProperties?: readonly PendingPropertyProposal[];
  /** GFM by default; auto preserves native merged and structured cells with HTML. */
  readonly tableSyntax?: MarkdownLayoutOptions['tableSyntax'];
  /** Include bounded embedded image attachments from this exact projected review view. */
  readonly includeImageAssets?: boolean;
};

/** Native immutable read records. These contain no live editor or export session. */
export interface ReviewDocumentView {
  readonly kind: DocumentView;
  readonly revision: RevisionToken;
  readonly layout: ExportSemanticLayout;
  readonly projection: MdxProjection;
  readonly media?: readonly MarkdownImageAsset[];
}

export type ReviewViewResult =
  { readonly success: true; readonly view: ReviewDocumentView } | NativeEditFailure;

function actualRevision(source: HeadlessDocumentView, expected: RevisionToken): RevisionToken {
  return { ...expected, revision: source.packageRevision() };
}

function stale(expected: RevisionToken, actual: RevisionToken): NativeEditFailure {
  return {
    success: false,
    code: 'stale-revision',
    error: 'The document changed while its review view was read.',
    expectedRevision: { ...expected },
    actualRevision: actual,
  };
}

/**
 * Read an explicit native review projection. Callers settle queued typing before creating
 * the revision token. This reader never flushes input, resolves proposals, or writes a part.
 */
export async function readNativeReviewView(
  source: HeadlessDocumentView,
  revision: RevisionToken,
  kind: DocumentView,
  options: ReviewViewOptions = {}
): Promise<ReviewViewResult> {
  const expected = Object.freeze({ ...revision });
  const before = actualRevision(source, expected);
  if (!sameRevision(expected, before)) return stale(expected, before);
  const pkg = source.currentPackage(),
    stored = readApplicationProposals(pkg, expected);
  if (!stored.success) return stored;
  const properties = projectPendingProperties(
    pkg,
    expected,
    [...stored.properties, ...(options.pendingProperties ?? [])],
    kind
  );
  if (!properties.success) return properties;
  const presence = projectApplicationPresence(properties.pkg, stored.presence, kind);
  if (!presence.success) return presence;
  const readPackage = withoutTaskCompletionAppearance(presence.pkg);
  const projectedSource =
    readPackage === pkg
      ? source
      : headlessViewOfStore(
          new TreePackageStore(readPackage, readPackage.parts.get(readPackage.mainDocumentPart)!)
        );
  const layoutSource =
    projectedSource === source
      ? source
      : {
          ...projectedSource,
          packageRevision: () => expected.revision,
        };
  const displayMode = kind === 'committed' ? 'original' : 'proposed';
  const {
    pendingProperties: _pendingProperties,
    tableSyntax,
    includeImageAssets,
    ...nativeOptions
  } = options;
  const opened = openDocumentForExport(layoutSource, {
    ...nativeOptions,
    displayMode,
    reuseAcrossRevisions: false,
  });
  if (!opened.ok)
    return {
      success: false,
      code: 'document-not-ready',
      error: `Unable to read the ${kind} document view: ${opened.reason}.`,
    };
  try {
    const layout = await opened.session.layout();
    const after = actualRevision(source, expected);
    if (!sameRevision(expected, after)) return stale(expected, after);
    const projected = projectMdx(presence.pkg, layout, expected, undefined, { tableSyntax });
    if (!projected.success) return projected;
    return {
      success: true,
      view: Object.freeze({
        kind,
        revision: expected,
        layout,
        projection: projected.projection,
        ...(includeImageAssets ? { media: nativeImageAssets(presence.pkg, layout) } : {}),
      }),
    };
  } catch (cause) {
    const after = actualRevision(source, expected);
    if (!sameRevision(expected, after)) return stale(expected, after);
    return {
      success: false,
      code: 'document-not-ready',
      error: `Unable to read the ${kind} document view: ${cause instanceof Error ? cause.message : String(cause)}`,
    };
  } finally {
    opened.session.dispose();
  }
}
