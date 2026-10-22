// Application-owned schema rules from the captured Sciobot validation engine.
// See fixtures/target-validation.json for source hashes and independent expected results.
export type ValidationIssue = {
  /** Stable machine code (parse ruleId, or a rule's configured code). */
  code: string;
  /** Human/agent-readable description, including fix guidance. */
  message: string;
  line?: number;
  column?: number;
};

export type ValidationResult = {
  ok: boolean;
  issues: ValidationIssue[];
};

/**
 * Cardinality of a child element (list, phases, slots…). `exactly` / `atLeast`
 * declare REQUIRED children; `atMost` declares an OPTIONAL child with a ceiling
 * (zero occurrences is fine).
 */
export type CountRule = {
  mode: 'exactly' | 'atLeast' | 'atMost';
  n: number;
};

/** When a repeated element must appear in a fixed order (e.g. the EUR triple). */
export type OrderedByRule = {
  /** Attribute whose value is checked against `sequence` per position. */
  attribute: string;
  /** Expected value at position i of the matched run (length should equal count.n). */
  sequence: readonly string[];
  /** The issue code to use for an out-of-order element (default: `<name>-order`). */
  orderCode?: string;
  /**
   * Message factory: (elementLabel, index, expectedValue, actualValue).
   * Defaults: `"<name> #${index+1} must have ${attribute}="${sequence[i]}" …"`.
   */
  orderMessage?: (
    label: string,
    index: number,
    expected: string,
    actual: string | undefined
  ) => string;
};

/** Declares a required child element of a container. */
export type ChildRule = {
  count: CountRule;
  orderedBy?: OrderedByRule;
  /** Issue-code overrides. Defaults: missing=`missing-<name>`,
   *  duplicate=`duplicate-<name>` (only when count is exactly 1),
   *  count=`<name>-count`, order=`<name>-order`. */
  codes?: { missing?: string; count?: string; order?: string };
  messages?: {
    missing?: string;
    /** Message factory for a wrong multiplicity (found ≠ expected). */
    count?: (found: number) => string;
  };
};

/** A compensating message; a function receives the offending value. */
export type IssueMessage = string | ((value: string) => string);

/** Attribute constraint applied to the element node carrying this rule. */
export type AttributeRule = {
  required?: boolean;
  /** When present, the literal string must be non-empty after trim. */
  nonEmpty?: boolean;
  /** When present, the literal string must match. */
  pattern?: RegExp;
  missingCode?: string;
  missingMessage?: IssueMessage;
  invalidCode?: string;
  invalidMessage?: IssueMessage;
};

function resolveMessage(message: IssueMessage | undefined, value: string): string | undefined {
  return typeof message === 'function' ? message(value) : message;
}

/** A custom predicate that flags the element node itself. */
export type NodeCheck = {
  code: string;
  message: string;
  /** Return true to raise an issue (i.e. the node FAILS the check). */
  test: (node: JsxNode) => boolean;
};

/** Full rule set for one container (keyed by MDX tag name in `containers`). */
export type ContainerRule = {
  /** Required children by MDX tag name. */
  children?: Record<string, ChildRule>;
  /** Attribute constraints on the container node itself. */
  attributes?: Record<string, AttributeRule>;
  /** Custom structural checks on the container node. */
  checks?: NodeCheck[];
};

/**
 * The declarative schema. `root` describes the document's top-level children;
 * `containers` describes every validated deep tag by name.
 */
export type MdxValidationSchema = {
  root: ContainerRule;
  containers: Record<string, ContainerRule>;
};

// ── AST types ─────────────────────────────────────────────────────────────

/** The shape of a JSX element within the parsed MDX tree. Exported so rule
 *  authors can write declarative predicates against it. */
export type JsxNode = {
  type: string;
  name?: string;
  value?: unknown;
  children?: JsxNode[];
  attributes?: MdxAttribute[];
  position?: {
    start?: { line?: number; column?: number };
    end?: { line?: number; column?: number };
  };
};

type MdxAttribute = {
  type: string;
  name?: string;
  value?: string | { type?: string; value?: string } | null;
};

function isJsx(node: JsxNode): boolean {
  return node.type === 'mdxJsxFlowElement' || node.type === 'mdxJsxTextElement';
}

function posOf(node: JsxNode): { line?: number; column?: number } {
  return {
    line: node.position?.start?.line,
    column: node.position?.start?.column,
  };
}

/**
 * Structural children of a container — JSX block children, or inline JSX
 * elements wrapped in paragraphs. remark-mdx promotes a tag to
 * `mdxJsxFlowElement` only when it contains block structure; a single-line
 * tag stays `mdxJsxTextElement` inside a paragraph. Both are real children
 * for shape validation.
 */
export function structuralChildren(node: JsxNode): JsxNode[] {
  const out: JsxNode[] = [];
  for (const child of node.children ?? []) {
    if (isJsx(child)) {
      out.push(child);
      continue;
    }
    if (child.type === 'paragraph') {
      for (const inline of child.children ?? []) {
        if (isJsx(inline)) out.push(inline);
      }
    }
  }
  return out;
}

/** Read an attribute's plain string literal value, or `undefined`. */
function attrValue(attrs: MdxAttribute[] | undefined, name: string): string | undefined {
  if (!attrs) return undefined;
  for (const attr of attrs) {
    if (attr.type !== 'mdxJsxAttribute' || attr.name !== name) continue;
    if (typeof attr.value === 'string') return attr.value;
    if (attr.value === null || attr.value === undefined) return '';
  }
  return undefined;
}

function defaultCodes(name: string, count: CountRule, explicit?: ChildRule['codes']) {
  // Hyphenate the tag name in generated codes so `recap_goals` → `duplicate-
  // recap-goals`, matching the underscore-free code convention.
  const slug = name.replace(/_/g, '-');
  const countCode =
    explicit?.count ??
    (count.mode === 'exactly' && count.n === 1 ? `duplicate-${slug}` : `${slug}-count`);
  return {
    missing: explicit?.missing ?? `missing-${slug}`,
    count: countCode,
    order: explicit?.order ?? `${slug}-order`,
  };
}

function checkChildCount(
  name: string,
  found: number,
  rule: ChildRule,
  issues: ValidationIssue[]
): void {
  const codes = defaultCodes(name, rule.count, rule.codes);

  if (found === 0) {
    if (rule.count.mode === 'atMost' || rule.count.n === 0) return;
    issues.push({
      code: codes.missing,
      message:
        rule.messages?.missing ??
        `Document must contain <${name}> (expected ${friendlyCount(rule.count)}).`,
    });
    return;
  }

  const countMatches =
    rule.count.mode === 'exactly'
      ? found === rule.count.n
      : rule.count.mode === 'atLeast'
        ? found >= rule.count.n
        : found <= rule.count.n;
  if (countMatches) return;

  issues.push({
    code: codes.count,
    message: rule.messages?.count?.(found) ?? expectedCountMessage(name, found, rule.count),
  });
}

function friendlyCount(count: CountRule): string {
  if (count.mode === 'exactly') return `exactly ${count.n}`;
  if (count.mode === 'atLeast') return `at least ${count.n}`;
  return `at most ${count.n}`;
}

function expectedCountMessage(name: string, found: number, count: CountRule): string {
  return `Expected ${friendlyCount(count)} <${name}>, found ${found}.`;
}

function checkChildOrder(
  name: string,
  matches: JsxNode[],
  rule: OrderedByRule,
  codes: ChildRule['codes'],
  issues: ValidationIssue[]
): void {
  if (!rule.sequence || rule.sequence.length === 0) return;
  const orderCode = codes?.order ?? `${name}-order`;
  for (let i = 0; i < Math.min(matches.length, rule.sequence.length); i++) {
    const node = matches[i]!;
    const actual = attrValue(node.attributes, rule.attribute);
    const expected = rule.sequence[i]!;
    if (actual !== expected) {
      issues.push({
        code: orderCode,
        message:
          rule.orderMessage?.(`<${name}>`, i, expected, actual) ??
          `<${name}> #${i + 1} must have ${rule.attribute}="${expected}" (got ${
            actual === undefined ? 'missing attribute' : `${rule.attribute}="${actual}"`
          }).`,
        ...posOf(node),
      });
    }
  }
}

function checkAttributes(
  node: JsxNode,
  rules: Record<string, AttributeRule> | undefined,
  issues: ValidationIssue[]
): void {
  if (!rules) return;
  const tag = `<${node.name ?? 'element'}>`;
  for (const [attrName, rule] of Object.entries(rules)) {
    const value = attrValue(node.attributes, attrName);
    const isMissing = value === undefined;

    if (rule.required && isMissing) {
      issues.push({
        code: rule.missingCode ?? 'missing-attribute',
        message: resolveMessage(rule.missingMessage, '') ?? `${tag} is missing ${attrName}="…".`,
        ...posOf(node),
      });
      continue;
    }
    if (isMissing) continue;

    const trimmed = value.trim();
    if (rule.nonEmpty && trimmed === '') {
      issues.push({
        code: rule.invalidCode ?? 'invalid-attribute',
        message:
          resolveMessage(rule.invalidMessage, '') ?? `${tag} has an empty ${attrName} attribute.`,
        ...posOf(node),
      });
      continue;
    }
    if (rule.pattern && !rule.pattern.test(trimmed)) {
      issues.push({
        code: rule.invalidCode ?? 'invalid-attribute',
        message:
          resolveMessage(rule.invalidMessage, value) ??
          `${tag} has an invalid ${attrName}="${value}".`,
        ...posOf(node),
      });
    }
  }
}

function checkNodeChecks(node: JsxNode, rule: ContainerRule, issues: ValidationIssue[]): void {
  for (const check of rule.checks ?? []) {
    if (check.test(node)) {
      issues.push({
        code: check.code,
        message: check.message,
        ...posOf(node),
      });
    }
  }
}

/** Validate the bounded AST returned by parseTargetMarkdown. This function performs no parse or writes. */
export function validateMdxTree(tree: JsxNode, schema: MdxValidationSchema): ValidationResult {
  const issues: ValidationIssue[] = [];
  walk(tree, schema.root, schema, issues);
  return { ok: issues.length === 0, issues };
}

function walk(
  parent: JsxNode,
  parentRule: ContainerRule,
  schema: MdxValidationSchema,
  issues: ValidationIssue[]
): void {
  checkNodeChecks(parent, parentRule, issues);
  checkAttributes(parent, parentRule.attributes, issues);

  const kids = structuralChildren(parent);
  for (const [name, childRule] of Object.entries(parentRule.children ?? {})) {
    const matches = kids.filter((k) => k.name === name);
    checkChildCount(name, matches.length, childRule, issues);
    if (childRule.orderedBy) {
      checkChildOrder(name, matches, childRule.orderedBy, childRule.codes, issues);
    }
    const childContainerRule = schema.containers[name];
    if (!childContainerRule) continue;
    for (const match of matches) {
      walk(match, childContainerRule, schema, issues);
    }
  }
}
