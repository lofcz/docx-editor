import type { MdxComponentStorage } from './contract';

/** The names Sciobot documents carry. Documents saved before the registry existed use them. */
export const SCIOBOT_STORAGE: MdxComponentStorage = Object.freeze({
  tagPrefix: 'scio:1:',
  namespace: 'urn:sciobot:docx:blocks:1',
  root: 'blocks',
  nodePrefix: 'scio_',
  reviewNodePrefix: 'scio-review_',
  labelTag: 'scio:label',
  adapterPrefix: 'scio',
  adapterNamespace: 'urn:sciobot:docx',
});
