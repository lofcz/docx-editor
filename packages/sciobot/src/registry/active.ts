import type { MdxComponentRegistry } from './contract';
import { sciobotMdxRegistry } from './sciobot';

let active: MdxComponentRegistry | undefined;

/**
 * Set the application's MDX components for this adapter instance.
 *
 * Call it once at startup, before an editor opens a document. Every adapter function that takes
 * no explicit registry reads this one, including the payload storage names. Until an application
 * configures one, the Sciobot registration applies, so existing documents keep opening.
 */
export function configureMdxRegistry(registry: MdxComponentRegistry): void {
  active = registry;
}

/** The configured registry, or the Sciobot registration when none is configured. */
export function activeMdxRegistry(): MdxComponentRegistry {
  return active ?? sciobotMdxRegistry();
}

/** Run `fn` with `registry` active, then restore the previous one. For tests and tools. */
export function withMdxRegistry<Result>(registry: MdxComponentRegistry, fn: () => Result): Result {
  const previous = active;
  active = registry;
  try {
    return fn();
  } finally {
    active = previous;
  }
}

/** The configured payload storage names. */
export function activeStorage(): MdxComponentRegistry['storage'] {
  return activeMdxRegistry().storage;
}
