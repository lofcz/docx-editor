import { remarkTargetMentions } from '../../src/projection/target-mentions';

// Mirrors the host Plate plugin's existing declaration. The private parser must not
// register a competing ambient mention node with different optional properties.
declare module 'mdast' {
  interface StaticPhrasingContentMap {
    mention: {
      children: { type: 'text'; value: string }[];
      type: 'mention';
      username: string;
      displayText?: string;
    };
  }
}

void remarkTargetMentions;
