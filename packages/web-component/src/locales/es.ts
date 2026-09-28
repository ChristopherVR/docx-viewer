import type { LocaleStrings } from './en';
import { esApp } from './es-app';
import { esRibbon } from './es-ribbon';

export const es: LocaleStrings = { ...esRibbon, ...esApp };
