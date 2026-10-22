// Development-only oracle. Execute the actual application reconciliation and Plate atomic diff policy.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const program = `
import { computeDiff } from '@platejs/diff';
import { lessonPlanDiffStrategy } from './src/components/lesson-plan/diff-strategy.ts';
import { reconcileListeningNodes } from './src/lib/reconcileListeningNodes.ts';
import { createEmptyAudioExercise } from './supabase/functions/_shared/activities/audio-exercise/empty.ts';
import { computeDialogueHash, dialogueHashTurns } from './supabase/functions/_shared/activities/audio-exercise/hash.ts';
import { assignMissingVoices } from './supabase/functions/_shared/voice/registry.ts';
const clone = value => JSON.parse(JSON.stringify(value));
const exercise = { ...createEmptyAudioExercise(), title:'At the market', intro:'Listen.', level:'A2',
  speakers:[{id:'s1',name:'Anna',gender:'female',voiceKey:'',voiceId:'',color:'violet',portraitUrl:'https://example.com/portrait.png'}],
  turns:[{id:'turn-1',speakerId:'s1',text:'Good morning!'}], tasks:[{id:'task-1',type:'open',question:'Who speaks?'}] };
assignMissingVoices(exercise.languageCode, exercise.speakers);
const seed = 7;
const hash = computeDialogueHash(dialogueHashTurns(exercise,text=>text),exercise.settings.model,exercise.languageCode,seed);
const original = {type:'listening',blockKey:'budget-keep',exercise,audioRef:{hash,seed},children:[{text:''}]};
const make = (name, edit, hydrated=false) => {
  const before = clone(original);
  if (hydrated) before.exercise.audio={url:'https://example.com/dialogue.mp3',hash,seed,durationMs:1000,turns:[{turnId:'turn-1',startMs:0,endMs:1000,words:[]}]};
  reconcileListeningNodes([before],[clone(before)]);
  const after = clone(before);
  after.blockKey='new-parser-key';
  delete after.audioRef; delete after.exercise.audio;
  after.exercise.tasks=[]; after.exercise.settings={model:'eleven_v3',allowTranscriptReveal:false};
  for (const speaker of after.exercise.speakers) { delete speaker.portraitUrl; speaker.voiceKey=''; speaker.voiceId=''; }
  after.exercise.turns[0].id='parsed-turn';
  edit(after);
  reconcileListeningNodes([after],[before]);
  const diff = computeDiff([before],[after], {
    granularity:'block',runScopeWordHints:true,ignoreProps:['id','listStart'],isInline:()=>false,
    getDeleteProps:()=>({change:'delete'}),getInsertProps:()=>({change:'insert'}),getUpdateProps:()=>({change:'update'}),getDiffStrategy:lessonPlanDiffStrategy,
  });
  return {name,before,after,rootChanges:diff.flatMap(node=>node.change?[node.change]:[])};
};
const cases = [
  make('same script with omitted derived fields',()=>{}),
  make('title change retains audio reference',node=>{node.exercise.title='At the shop';}),
  make('script change invalidates audio reference',node=>{node.exercise.turns[0].text='Good evening!';}),
  make('hydrated unchanged audio',()=>{},true),
  make('hydrated script change',node=>{node.exercise.turns[0].text='Good evening!';},true),
];
process.stdout.write(JSON.stringify(cases));
`;
const cases = JSON.parse(execFileSync('bun', ['--eval', program], { cwd: app, encoding: 'utf8' }));
const paths = [
  'src/components/lesson-plan/diff-strategy.ts',
  'src/lib/reconcileListeningNodes.ts',
  'src/lib/ai-review-diff.ts',
  'supabase/functions/_shared/activities/audio-exercise/empty.ts',
  'supabase/functions/_shared/activities/audio-exercise/hash.ts',
  'supabase/functions/_shared/activities/audio-exercise/merge.ts',
  'supabase/functions/_shared/activities/audio-exercise/normalize.ts',
  'supabase/functions/_shared/activities/audio-exercise/portrait-catalog.ts',
  'supabase/functions/_shared/activities/audio-exercise/questions.ts',
  'supabase/functions/_shared/activities/audio-exercise/types.ts',
  'supabase/functions/_shared/voice/registry.ts',
  'supabase/functions/_shared/voice/taxonomy.ts',
  'supabase/functions/_shared/voice/voices.generated.ts',
  'node_modules/@platejs/diff/dist/index.js',
];
const fixture = {
  applicationRevision: execFileSync('git', ['rev-parse', 'HEAD'], {
    cwd: app,
    encoding: 'utf8',
  }).trim(),
  sources: Object.fromEntries(
    paths.map((path) => [
      path,
      createHash('sha256')
        .update(readFileSync(resolve(app, path)))
        .digest('hex'),
    ])
  ),
  cases,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/atomic-assets.json'),
  `${JSON.stringify(fixture, null, 2)}\n`
);
process.stdout.write(`${cases.length} application atomic asset cases captured.\n`);
