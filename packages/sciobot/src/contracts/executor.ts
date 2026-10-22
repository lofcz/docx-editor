/** Private application ports. These are not a public document automation API. */
export interface ReadInput {
  offset?: number;
  limit?: number;
}

export interface ReadResult {
  content: string;
  startLine: number;
  endLine: number;
  totalLines: number;
  truncated: boolean;
}

export interface SearchInput {
  pattern: string;
  isRegex: boolean;
  maxResults: number;
}

export interface SearchMatch {
  lineNumber: number;
  content: string;
  contextBefore: string | null;
  contextAfter: string | null;
}

export interface SearchResult {
  readonly matches: readonly SearchMatch[];
  totalMatches: number;
  truncated: boolean;
}

export interface SearchError {
  readonly matches: readonly [];
  totalMatches: 0;
  error: string;
}

export interface EditInput {
  isSuggestion?: boolean;
  oldString: string;
  newString: string;
  replaceAll: boolean;
  expectedReplacements?: number;
}

export interface SetInput {
  isSuggestion?: boolean;
  markdown: string;
}

export type EditResult = { success: true; message: string } | { success: false; error: string };
export type SetResult = EditResult;

export interface FileExecutor {
  read(input: ReadInput): ReadResult;
  search(input: SearchInput): SearchResult | SearchError;
  edit?(input: EditInput): EditResult;
  set?(input: SetInput): SetResult;
}

interface WireScope {
  readonly path: string;
  readonly variantId?: string;
}

export interface ReadDocumentInput extends WireScope {
  readonly offset?: number;
  readonly limit?: number;
}

export interface SearchDocumentInput extends WireScope {
  readonly pattern: string;
  readonly is_regex?: boolean;
  readonly max_results?: number;
}

export interface EditDocumentInput extends WireScope {
  /** This spelling belongs to the existing agent protocol. */
  readonly is_suggesion?: boolean;
  readonly old_string: string;
  readonly new_string: string;
  readonly replace_all?: boolean;
  readonly expected_replacements?: number;
}

export interface SetDocumentInput extends WireScope {
  readonly is_suggesion?: boolean;
  readonly markdown: string;
}

export type DocumentToolRequest =
  | { readonly tool: 'read_document'; readonly input: ReadDocumentInput }
  | { readonly tool: 'search_document'; readonly input: SearchDocumentInput }
  | { readonly tool: 'edit_document'; readonly input: EditDocumentInput }
  | { readonly tool: 'set_document'; readonly input: SetDocumentInput };
