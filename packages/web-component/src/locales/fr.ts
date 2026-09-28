import type { LocaleStrings } from './en';
import { frApp } from './fr-app';
import { frRibbon } from './fr-ribbon';

export const fr: LocaleStrings = { ...frRibbon, ...frApp };
