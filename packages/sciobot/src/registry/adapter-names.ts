import { activeStorage } from './storage';

// The adapter's own content-control tags and custom XML namespaces. Each derives from the
// configured storage, so another embedder's documents never carry these Sciobot names.

/** The tag of a checklist item's checkbox control. */
export const checklistTag = (): string => `${activeStorage().adapterPrefix}:list-checkbox`;

/** The prefix of a temporary control that reserves an inline image while it loads. */
export const imageReservationPrefix = (): string =>
  `${activeStorage().adapterPrefix}:image-reservation:`;

/** The prefix every adapter-owned control tag starts with. */
export const adapterTagPrefix = (): string => `${activeStorage().adapterPrefix}:`;

/** The namespace of list-continuity state kept on paragraphs. */
export const listStateNamespace = (): string => `${activeStorage().adapterNamespace}:lists:1`;

/** The namespace of review group markers and their record store. */
export const reviewGroupNamespace = (): string =>
  `${activeStorage().adapterNamespace}:review-groups:1`;
