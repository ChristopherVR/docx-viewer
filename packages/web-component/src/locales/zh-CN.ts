import type { LocaleStrings } from './en';
import { zhCNApp } from './zh-CN-app';
import { zhCNRibbon } from './zh-CN-ribbon';

export const zhCN: LocaleStrings = { ...zhCNRibbon, ...zhCNApp };
