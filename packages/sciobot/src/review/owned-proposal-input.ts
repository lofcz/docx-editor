import type { RevisionAddress } from '@docx-editor.dev/core/store';
import { isBlockId } from '../blocks/payload-schema';
import { GROUP_ID_LIMIT } from './group-markers';
import type { ReviewGroupPropertyMember } from './groups';
import { propertyAddressKey } from './property-fallback';

const object = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const keys = (value: object, allowed: readonly string[]) =>
  Object.keys(value).every((key) => allowed.includes(key));
const id = (value: unknown): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= GROUP_ID_LIMIT;
const partName = (value: unknown): value is string =>
  id(value) && value.startsWith('/') && !value.split('/').includes('..');
const propertyKey = (member: ReviewGroupPropertyMember) =>
  JSON.stringify([member.proposalId, member.blockId]);

function address(value: unknown): value is RevisionAddress {
  return (
    object(value) &&
    id(value.id) &&
    typeof value.author === 'string' &&
    value.author.length <= GROUP_ID_LIMIT &&
    (!Object.hasOwn(value, 'date') ||
      (typeof value.date === 'string' && value.date.length <= GROUP_ID_LIMIT)) &&
    keys(value, ['id', 'author', 'date'])
  );
}

export function ownedProposalSelectionProblem(value: unknown): string | null {
  if (!Array.isArray(value)) return 'Owned proposal selection must be an array.';
  const groups = new Set<string>();
  for (const group of value) {
    if (
      !object(group) ||
      !id(group.groupId) ||
      groups.has(group.groupId) ||
      !Array.isArray(group.native) ||
      !Array.isArray(group.properties) ||
      !keys(group, ['groupId', 'native', 'properties'])
    )
      return 'Owned groups need unique identities and exact receipt fields.';
    groups.add(group.groupId);
    const native = new Set<string>(),
      properties = new Set<string>();
    for (const member of group.native) {
      if (
        !object(member) ||
        !partName(member.partName) ||
        !Array.isArray(member.addresses) ||
        !member.addresses.length ||
        !keys(member, ['partName', 'addresses'])
      )
        return 'Invalid owned native receipt.';
      for (const revision of member.addresses) {
        if (!address(revision)) return 'Invalid owned native revision address.';
        const key = propertyAddressKey(member.partName, revision);
        if (native.has(key)) return 'Duplicate owned native revision receipt.';
        native.add(key);
      }
    }
    for (const member of group.properties) {
      if (
        !object(member) ||
        !id(member.proposalId) ||
        !isBlockId(member.blockId) ||
        !keys(member, ['proposalId', 'blockId'])
      )
        return 'Invalid owned application proposal receipt.';
      const key = propertyKey(member as unknown as ReviewGroupPropertyMember);
      if (properties.has(key)) return 'Duplicate owned application proposal receipt.';
      properties.add(key);
    }
    if (!native.size && !properties.size) return 'An owned group needs a proposal receipt.';
  }
  return null;
}
