import type { MdxComponentStorage } from './contract';

// Storage names have no runtime dependency, so light entry points (the collaboration protocol)
// can read the configured names without loading the component codecs or the editor runtime.

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

let configured: MdxComponentStorage | undefined;

/** Set the storage names. `configureMdxRegistry` calls it; nothing else should. */
export function setActiveStorage(storage: MdxComponentStorage | undefined): void {
  configured = storage;
}

/** The configured payload storage names, or the Sciobot names when none are configured. */
export function activeStorage(): MdxComponentStorage {
  return configured ?? SCIOBOT_STORAGE;
}
