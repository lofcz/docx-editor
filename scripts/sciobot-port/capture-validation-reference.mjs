// Development-only oracle. Execute the application's validator and feedback functions directly.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const read = (path) => readFileSync(resolve(app, path), 'utf8');
const valid = read('src/lib/lesson-plan-validation.test.ts').match(
  /const VALID_DOC = `([\s\S]*?)`;/
)?.[1];
if (!valid) throw new Error('Missing independent valid lesson fixture.');
const cases = [
  { name: 'valid complete lesson', input: valid },
  { name: 'generic document', input: '# Replacement\n\nOrdinary text.' },
  { name: 'empty document', input: '' },
  { name: 'whitespace only', input: ' \n\t' },
  { name: 'malformed unclosed tag', input: '<activity>Broken' },
  { name: 'malformed stray closing tag', input: valid + '\n</story>\n' },
  { name: 'invalid expression', input: 'Text {broken(}' },
  { name: 'missing title', input: valid.replace(/<lesson_title[\s\S]*?<\/lesson_title>\n\n/, '') },
  { name: 'duplicate title', input: '<lesson_title>Extra</lesson_title>\n\n' + valid },
  {
    name: 'missing recap',
    input: valid.replace(/<preparation_recap>[\s\S]*?<\/preparation_recap>\n\n/, ''),
  },
  { name: 'missing goals', input: valid.replace(/<recap_goals>[\s\S]*?<\/recap_goals>/, '') },
  {
    name: 'wrong phase order',
    input: valid.replace('<phase type="evocation">', '<phase type="reflection">'),
  },
  {
    name: 'wrong recap phase order',
    input: valid.replace('<recap_phase type="evocation">', '<recap_phase type="reflection">'),
  },
  { name: 'empty phase', input: valid.replace(/<activity name="Úvod"[\s\S]*?<\/activity>\n/, '') },
  {
    name: 'zero minute continuation',
    input: valid.replace(
      /<phase type="evocation">[\s\S]*?<\/phase>/,
      '<phase type="evocation" duration="0"></phase>'
    ),
  },
  { name: 'missing activity name', input: valid.replace('name="Úvod"', '') },
  { name: 'empty activity name', input: valid.replace('name="Úvod"', 'name="  "') },
  { name: 'missing activity duration', input: valid.replace('duration="10"', '') },
  { name: 'zero activity duration', input: valid.replace('duration="10"', 'duration="0"') },
  { name: 'fractional activity duration', input: valid.replace('duration="10"', 'duration="1.5"') },
  {
    name: 'empty activity',
    input: valid.replace(
      /<activity name="Úvod"[\s\S]*?<\/activity>/,
      '<activity name="Úvod" duration="10">\n</activity>'
    ),
  },
  {
    name: 'duplicate adaptations',
    input: valid.replace(
      '<recap_goals>',
      '<recap_adaptations>A</recap_adaptations>\n\n<recap_adaptations>B</recap_adaptations>\n\n<recap_goals>'
    ),
  },
  {
    name: 'valid summary visual',
    input: valid.replace(
      '</preparation_recap>',
      '</preparation_recap>\n\n<lesson_visual kind="summary" src="asset.png" />'
    ),
  },
  {
    name: 'misplaced summary visual',
    input: '<lesson_visual kind="summary" src="asset.png" />\n\n' + valid,
  },
  {
    name: 'duplicate summary visual',
    input: valid.replace(
      '</preparation_recap>',
      '</preparation_recap>\n\n<lesson_visual kind="summary" src="one.png" />\n\n<lesson_visual kind="summary" src="two.png" />'
    ),
  },
  {
    name: 'visual missing attributes',
    input: valid.replace('</preparation_recap>', '</preparation_recap>\n\n<lesson_visual />'),
  },
  {
    name: 'visual invalid attributes',
    input: valid.replace(
      '</preparation_recap>',
      '</preparation_recap>\n\n<lesson_visual kind="wrong" src="  " />'
    ),
  },
  {
    name: 'misplaced phase visual',
    input: valid.replace(
      '<activity name="Úvod"',
      '<lesson_visual kind="explainer" src="asset.png" />\n\n<activity name="Úvod"'
    ),
  },
  {
    name: 'inline JSX children',
    input: valid.replace(
      /<activity name="Úvod"[\s\S]*?<\/activity>/,
      '<activity name="Úvod" duration="10">Inline body.</activity>'
    ),
  },
  {
    name: 'inert expression attributes',
    input: valid.replace('name="Úvod"', 'name={globalThis.__validationProbe()}'),
  },
  {
    name: 'unknown tags stay tolerated',
    input: valid.replace(
      'Rozdělení uhlovodíků.',
      '<future-widget expr={globalThis.__validationProbe()}>Text.</future-widget>'
    ),
  },
  {
    name: 'math braces and line breaks',
    input: valid.replace(
      'Rozdělení uhlovodíků.',
      'Vzorec $C_8H_{16}$ a $$\\frac{a}{b}$$.<br>Text.'
    ),
  },
];
const tagCases = [
  { oldString: '    </callout>\n  </activity>', newString: '    </callout>' },
  { oldString: '<story>\n  first sentence', newString: '  first sentence' },
  { oldString: 'plain paragraph', newString: '<callout>\n  tip' },
  { oldString: 'plain paragraph', newString: 'text\n</phase>' },
  { oldString: '<source title="a" url="b" type="verified" />', newString: 'replacement text' },
  { oldString: '<story>\nbody\n</story>', newString: '' },
  { oldString: '<callout icon="a>b" meta={x > 1}>\nbody\n</callout>\n</activity>', newString: '' },
  {
    oldString: '</solution>\n<callout id="x">body</callout>\n</activity>',
    newString: '</solution>\n</activity>',
  },
  { oldString: '<a><a>', newString: '<a>' },
];
const output = JSON.parse(
  execFileSync(
    'bun',
    [
      '--eval',
      `
import { validateLessonPlanSyntax, validateLessonPlanMarkdown, formatEditRejection, formatSchemaWarning } from './src/lib/lesson-plan-validation.ts';
import { analyzeEditTagBalance } from './src/lib/agent-edit-feedback.ts';
import { normalizeMdxLineBreaks } from './supabase/functions/_shared/normalize-mdx-line-breaks.ts';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkMath from 'remark-math';
import remarkMdx from 'remark-mdx';
const cases = ${JSON.stringify(cases)}.map(({name,input}) => {
  const syntax = validateLessonPlanSyntax(input), lesson = validateLessonPlanMarkdown(input);
  const markdown = normalizeMdxLineBreaks(input);
  // Preserve primary parser point positions which the facade's range-only lookup omits.
  let locatedSyntax = syntax;
  if (!syntax.ok && syntax.issues[0]?.code !== 'empty-document') {
    try { unified().use(remarkParse).use(remarkMath).use(remarkMdx).parse(markdown); }
    catch (error) { locatedSyntax = { ...syntax, issues: syntax.issues.map((issue) => ({...issue,
      ...(issue.line === undefined && Number.isSafeInteger(error.line) ? {line:error.line} : {}),
      ...(issue.column === undefined && Number.isSafeInteger(error.column) ? {column:error.column} : {}),
    })) }; }
  }
  return {name,input,syntax,locatedSyntax,lesson,
    ...(!syntax.ok ? {rejection:formatEditRejection(syntax.issues,markdown)} :
       !lesson.ok ? {rejection:formatEditRejection(lesson.issues,markdown,undefined,'schema'),warning:formatSchemaWarning(lesson.issues)} : {}),
    ...(!syntax.ok ? {locatedRejection:formatEditRejection(locatedSyntax.issues,markdown)} :
       !lesson.ok ? {locatedRejection:formatEditRejection(lesson.issues,markdown,undefined,'schema')} : {}),
  };
});
const tagCases = ${JSON.stringify(tagCases)}.map((input)=>({...input,hints:analyzeEditTagBalance(input.oldString,input.newString)}));
process.stdout.write(JSON.stringify({cases,tagCases}));
`,
    ],
    { cwd: app, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }
  )
);
const files = [
  'src/lib/lesson-plan-validation.ts',
  'src/lib/lesson-plan-validation.test.ts',
  'src/lib/agent-edit-feedback.ts',
  'src/lib/agent-edit-feedback.test.ts',
  'supabase/functions/ai-material-draft/validate-mdx.ts',
  'supabase/functions/ai-material-draft/lesson-plan-schema.ts',
  'supabase/functions/_shared/normalize-mdx-line-breaks.ts',
];
const artifact = {
  schemaVersion: 1,
  source: {
    revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: app, encoding: 'utf8' }).trim(),
    files: Object.fromEntries(
      files.map((file) => [
        file,
        createHash('sha256')
          .update(readFileSync(resolve(app, file)))
          .digest('hex'),
      ])
    ),
    scope:
      'Application-owned syntax/schema validators and exact feedback; no native implementation is used as the oracle.',
  },
  ...output,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/target-validation.json'),
  JSON.stringify(artifact, null, 2) + '\n'
);
process.stdout.write(
  `Captured ${output.cases.length} syntax/schema cases and ${output.tagCases.length} tag-balance cases.\n`
);
