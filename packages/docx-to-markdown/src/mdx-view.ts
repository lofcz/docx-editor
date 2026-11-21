/** Display and register MDX components in native Word documents. @packageDocumentation @public */
export * from './mdx/registry.ts';
export {
  isMdxDocument,
  readMdxComponents,
  type MdxDocumentComponent,
} from './mdx/document-reader.ts';
export {
  configureMdxDisclosures,
  mountMdxDisclosures,
  type MdxDisclosureOptions,
} from './mdx/document-disclosures.ts';
export {
  mountMdxComponents,
  type MdxComponentRenderer,
  type MdxComponentView,
} from './mdx/document-views.ts';
