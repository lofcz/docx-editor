// Backspace and Delete remove one user-perceived character. A deletion of one UTF-16 unit
// splits a surrogate pair, which the store refuses, so the key did nothing after an emoji.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test, afterEach } from 'bun:test';
import { mount, paragraph, putCaret } from './paginated-surface-fixtures.ts';

afterEach(() => {
  document.getSelection()?.removeAllRanges();
});

describe('grapheme deletion', () => {
  for (const emoji of ['🚀', '🏳️‍🌈', '👍🏽', 'é']) {
    test(`backspace removes ${emoji} whole`, () => {
      const { surface } = mount(paragraph(`A ${emoji}`));
      putCaret(surface, 2 + emoji.length);
      surface.deleteBackward();
      expect(surface.session.bodyText()).toBe('A ');
      expect(surface.state().selection.head.offset).toBe(2);
      surface.undo();
      expect(surface.session.bodyText()).toBe(`A ${emoji}`);
    });

    test(`delete removes ${emoji} whole`, () => {
      const { surface } = mount(paragraph(`${emoji} B`));
      putCaret(surface, 0);
      surface.deleteForward();
      expect(surface.session.bodyText()).toBe(' B');
      expect(surface.state().selection.head.offset).toBe(0);
    });
  }

  test('backspace still removes one letter', () => {
    const { surface } = mount(paragraph('abc'));
    putCaret(surface, 3);
    surface.deleteBackward();
    expect(surface.session.bodyText()).toBe('ab');
  });
});
