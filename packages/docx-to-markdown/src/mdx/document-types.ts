/** The structural Markdown tree used at the Word conversion boundary. */
export interface DocumentNode {
  type: string;
  name?: string | null;
  value?: string;
  children?: DocumentNode[];
  [key: string]: unknown;
}
