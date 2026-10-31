import type { MdxComponentStorage } from './contract';
import { SCIOBOT_STORAGE } from './sciobot-storage';

export { SCIOBOT_STORAGE };

// Storage names have no runtime dependency, so light entry points (the collaboration protocol)
// can read the configured names without loading the component codecs or the editor runtime.

let configured: MdxComponentStorage | undefined;

/** Set the storage names. `configureMdxRegistry` calls it; nothing else should. */
export function setActiveStorage(storage: MdxComponentStorage | undefined): void {
  configured = storage;
}

/** The configured payload storage names, or the Sciobot names when none are configured. */
export function activeStorage(): MdxComponentStorage {
  return configured ?? SCIOBOT_STORAGE;
}
