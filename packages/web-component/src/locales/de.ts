import type { LocaleStrings } from './en';
import { deApp } from './de-app';
import { deRibbon } from './de-ribbon';

export const de: LocaleStrings = { ...deRibbon, ...deApp };
