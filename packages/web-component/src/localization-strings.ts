import { de } from './locales/de';
import { en } from './locales/en';
import { es } from './locales/es';
import { fr } from './locales/fr';
import { zhCN } from './locales/zh-CN';

export type { LocaleStrings, LocalizationKey } from './locales/en';

/** Every display locale, keyed by its canonical code. `en` defines the key set. */
export const strings = { en, fr, de, es, 'zh-CN': zhCN } as const;

/** Canonical codes of the supported display locales, in menu order. */
export const EDITOR_LOCALES = ['en', 'fr', 'de', 'es', 'zh-CN'] as const;
