// Application-owned tag-balance hints. See fixtures/target-validation.json.
interface TagBalance {
  unclosedOpens: string[];
  orphanCloses: string[];
}

/** Linear best-effort tokenization. The MDX parser owns syntax acceptance. */
function* tagsOf(text: string) {
  let next = 0;
  while (next < text.length) {
    const start = text.indexOf('<', next);
    if (start < 0) return;
    let at = start + 1;
    const close = text[at] === '/';
    if (close) at++;
    const first = at;
    if (!/[A-Za-z]/.test(text[at] ?? '')) {
      next = at + 1;
      continue;
    }
    while (at < text.length && /[A-Za-z0-9_-]/.test(text[at]!)) at++;
    const name = text.slice(first, at);
    let quote: string | null = null,
      brace = false,
      complete = false;
    while (at < text.length) {
      const character = text[at]!;
      if (quote) {
        if (character === quote) quote = null;
      } else if (brace) {
        if (character === '}') brace = false;
      } else if (character === '"' || character === "'") quote = character;
      else if (character === '{') brace = true;
      else if (character === '>') {
        yield { name, close, selfClosing: text[at - 1] === '/' };
        next = at + 1;
        complete = true;
        break;
      }
      at++;
    }
    if (!complete) return;
  }
}

function computeTagBalance(text: string): TagBalance {
  const opens: { name: string; alive: boolean }[] = [];
  const byName = new Map<string, number[]>();
  const orphanCloses: string[] = [];
  for (const tag of tagsOf(text)) {
    if (tag.selfClosing) continue;
    if (tag.close) {
      const index = byName.get(tag.name)?.pop();
      if (index === undefined) orphanCloses.push(tag.name);
      else opens[index]!.alive = false;
    } else {
      const positions = byName.get(tag.name) ?? [];
      positions.push(opens.length);
      byName.set(tag.name, positions);
      opens.push({ name: tag.name, alive: true });
    }
  }
  return {
    unclosedOpens: opens.filter((item) => item.alive).map((item) => item.name),
    orphanCloses,
  };
}

/** Multiset difference a \ b. */
function* diffCounts(a: string[], b: string[]) {
  const counts = new Map<string, number>();
  for (const name of a) counts.set(name, (counts.get(name) ?? 0) + 1);
  for (const name of b) counts.set(name, (counts.get(name) ?? 0) - 1);
  for (const [name, n] of counts) {
    for (let i = 0; i < n; i++) yield name;
  }
}

/**
 * When an exact-string edit gets REJECTED for unbalanced MDX, pinpoint the
 * tag pair the replacement broke.
 *
 * An edit is structurally neutral when new_string carries the SAME set of
 * boundary-crossing tags as old_string: every close tag in old_string whose
 * open lies outside the edited text must be re-emitted in new_string (and
 * vice versa for opens). Any difference is reported as a precise hint.
 */
export function analyzeEditTagBalance(oldString: string, newString: string): string[] {
  const oldBal = computeTagBalance(oldString);
  const newBal = computeTagBalance(newString);

  const hints: string[] = [];
  let omitted = false;
  const append = (message: () => string) => {
    if (hints.length < 64) hints.push(message());
    else omitted = true;
  };

  for (const name of diffCounts(oldBal.orphanCloses, newBal.orphanCloses)) {
    append(
      () =>
        `Your old_string contains </${name}> which closes a <${name}> that opens OUTSIDE the edited text, but your new_string does not re-emit it — the document would lose that closing tag. Either re-emit </${name}> in new_string, or shorten old_string so it does not include the closing tag.`
    );
  }
  for (const name of diffCounts(newBal.orphanCloses, oldBal.orphanCloses)) {
    append(
      () =>
        `Your new_string adds </${name}> with no matching <${name}> — the document would gain a stray closing tag. Remove it or add the opening tag.`
    );
  }
  for (const name of diffCounts(oldBal.unclosedOpens, newBal.unclosedOpens)) {
    append(
      () =>
        `Your old_string contains <${name}> whose closing </${name}> lies OUTSIDE the edited text, but your new_string does not re-emit it — the document would keep an orphaned </${name}>. Either re-emit <${name}> in new_string, or extend old_string to include the closing tag as well.`
    );
  }
  for (const name of diffCounts(newBal.unclosedOpens, oldBal.unclosedOpens)) {
    append(
      () =>
        `Your new_string opens <${name}> but never closes it, and old_string did not leave a matching </${name}> in place. Close the tag inside new_string.`
    );
  }

  if (omitted) hints.push('Additional tag-boundary hints were omitted after the first 64.');
  return hints;
}
