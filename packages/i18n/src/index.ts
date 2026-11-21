/**
 * Shared locale data, types, and runtime helpers for the @docx-editor.dev
 * editor adapters.
 *
 * Import everything from the package root. `sideEffects: false` lets
 * consumer bundlers tree-shake unused locales.
 *
 * ```ts
 * import {
 *   en, de, pl, tr, he, ptBR, zhCN,    // typed locale data
 *   locales,                            // record keyed by BCP-47 tag
 *   deepMerge, createT,                 // build a t() for custom hosts
 *   type LocaleStrings,                 // shape of `en` (source of truth)
 *   type Translations,                  // shape of a community partial
 *   type TranslationKey,                // every valid `t()` key
 *   type LocaleCode,                    // 'en' | 'de' | 'pt-BR' | ...
 * } from '@docx-editor.dev/i18n';
 * ```
 *
 * The React and Vue adapters wrap `createT` in framework-native bindings
 * (`useTranslation`, `LocaleProvider`, etc.); use those for app code.
 * Reach for `createT` directly when building a non-React/Vue host.
 *
 * @packageDocumentation
 * @public
 */

// ─── GENERATED START — `bun run i18n:codegen` ───
// DO NOT EDIT — this block is rewritten from the on-disk
// `packages/i18n/*.json` filenames whenever `bun run i18n:codegen`
// runs (and `bun run i18n:new <lang>` runs it automatically).
// `bun run i18n:validate` fails CI if hand-edits drift from the
// JSON files. Edit the JSON, not this block.

import enJson from '../en.json';
import arJson from '../ar.json';
import deJson from '../de.json';
import esJson from '../es.json';
import frJson from '../fr.json';
import heJson from '../he.json';
import hiJson from '../hi.json';
import idJson from '../id.json';
import itJson from '../it.json';
import jaJson from '../ja.json';
import plJson from '../pl.json';
import ptBRJson from '../pt-BR.json';
import ruJson from '../ru.json';
import trJson from '../tr.json';
import zhCNJson from '../zh-CN.json';

/**
 * Full locale string set, auto-derived from `en.json` (the source of truth).
 * Every other locale is a `PartialLocaleStrings` against this shape.
 *
 * @public
 */
export type LocaleStrings = typeof enJson;

/**
 * Every locale code shipped from this package. Pass to `locales[code]`
 * for runtime lookup; assign to `_lang` to drive `Intl.PluralRules`.
 *
 * Custom codes are accepted at runtime ({@link PartialLocaleStrings._lang}
 * widens to any string), but the shipped union is the IDE-completion list.
 *
 * @public
 */
// prettier-ignore
export type LocaleCode = 'en' | 'ar' | 'de' | 'es' | 'fr' | 'he' | 'hi' | 'id' | 'it' | 'ja' | 'pl' | 'pt-BR' | 'ru' | 'tr' | 'zh-CN';

/** English (`en`) — the source of truth, 100% covered. @public */
export const en: LocaleStrings = enJson;

/** Arabic (`ar`). Community-maintained; null leaves fall back to English. @public */
export const ar: PartialLocaleStrings = arJson;

/** German (`de`). Community-maintained; null leaves fall back to English. @public */
export const de: PartialLocaleStrings = deJson;

/** Spanish (`es`). Community-maintained; null leaves fall back to English. @public */
export const es: PartialLocaleStrings = esJson;

/** French (`fr`). Community-maintained; null leaves fall back to English. @public */
export const fr: PartialLocaleStrings = frJson;

/** Hebrew (`he`). Community-maintained; null leaves fall back to English. @public */
export const he: PartialLocaleStrings = heJson;

/** Hindi (`hi`). Community-maintained; null leaves fall back to English. @public */
export const hi: PartialLocaleStrings = hiJson;

/** Indonesian (`id`). Community-maintained; null leaves fall back to English. @public */
export const id: PartialLocaleStrings = idJson;

/** Italian (`it`). Community-maintained; null leaves fall back to English. @public */
export const it: PartialLocaleStrings = itJson;

/** Japanese (`ja`). Community-maintained; null leaves fall back to English. @public */
export const ja: PartialLocaleStrings = jaJson;

/** Polish (`pl`). Community-maintained; null leaves fall back to English. @public */
export const pl: PartialLocaleStrings = plJson;

/** Portuguese (Brazil) (`pt-BR`). Community-maintained; null leaves fall back to English. @public */
export const ptBR: PartialLocaleStrings = ptBRJson;

/** Russian (`ru`). Community-maintained; null leaves fall back to English. @public */
export const ru: PartialLocaleStrings = ruJson;

/** Turkish (`tr`). Community-maintained; null leaves fall back to English. @public */
export const tr: PartialLocaleStrings = trJson;

/** Simplified Chinese (`zh-CN`). Community-maintained; null leaves fall back to English. @public */
export const zhCN: PartialLocaleStrings = zhCNJson;

/**
 * Every shipped locale, keyed by BCP-47 tag. Use for runtime locale
 * pickers and "look up the locale matching this user preference" code:
 *
 * ```ts
 * <DocxEditor i18n={locales[userLocale]} />
 * ```
 *
 * Importing `locales` defeats the per-locale tree-shake — the bundler
 * sees a static reference to every locale. If you only need one or two,
 * import them by name (`import { en, de } from '...'`) instead.
 *
 * @public
 */
export const locales: Record<LocaleCode, PartialLocaleStrings> = {
  en,
  ar,
  de,
  es,
  fr,
  he,
  hi,
  id,
  it,
  ja,
  pl,
  'pt-BR': ptBR,
  ru,
  tr,
  'zh-CN': zhCN,
};
// ─── GENERATED END ───

/**
 * Recursive Partial that allows `null` at leaves to signal "not yet
 * translated, fall back to English." Community translations use this shape;
 * `bun run i18n:fix` keeps every locale aligned to `en.json` with `null`
 * placeholders for missing keys.
 *
 * @public
 */
export type DeepPartial<T> = {
  [K in keyof T]?: T[K] extends object ? DeepPartial<T[K]> : T[K] | null;
};

/**
 * Partial locale strings — what consumers pass to the editor's `i18n` prop.
 * Missing keys fall back to English. Optional `_lang` carries the BCP-47
 * tag used by `Intl.PluralRules`; shipped codes autocomplete but custom
 * strings are accepted.
 *
 * @public
 */
export type PartialLocaleStrings = DeepPartial<LocaleStrings> & {
  // `string & {}` keeps the literal-union completions for `LocaleCode`
  // while still accepting arbitrary strings at runtime.
  _lang?: LocaleCode | (string & {});
  /**
   * The writing direction of the interface strings. `'rtl'` mirrors the editor chrome, as
   * the shipped Arabic (`ar`) and Hebrew (`he`) catalogs do. Without it, the direction
   * follows `_lang` (see {@link localeDirection}).
   */
  _dir?: 'ltr' | 'rtl' | (string & {});
};

/**
 * Alias for `PartialLocaleStrings`. Prefer this name when typing the
 * consumer-facing `i18n` prop or function parameter.
 *
 * @public
 */
export type Translations = PartialLocaleStrings;

type DotPath<T, Prefix extends string = ''> = {
  [K in keyof T & string]: T[K] extends Record<string, unknown>
    ? DotPath<T[K], `${Prefix}${K}.`>
    : `${Prefix}${K}`;
}[keyof T & string];

/**
 * Every valid dot-notation key into `LocaleStrings`, e.g. `'formattingBar.bold'`
 * or `'navigation.find.counter'`. Pass to `t(key, vars?)` for
 * compile-time-checked translation lookup.
 *
 * @public
 */
export type TranslationKey = DotPath<LocaleStrings>;

// ─────────────────────────────────────────────────────────────────────────────
// Runtime helpers — used by the React and Vue adapters; exported here so
// non-React/Vue hosts can build a typed `t()` without re-implementing the
// merge + ICU format logic.
// ─────────────────────────────────────────────────────────────────────────────

export {
  createT,
  deepMerge,
  localizeDisabledReason,
  catalogDirection,
  chromeDirection,
  localeDirection,
  isApplePlatform,
  platformShortcut,
  type TFunction,
} from './translation-helpers';
