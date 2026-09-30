import { zhColumns } from './columns';
import { zhMultilevelList } from './multilevel-list';
import { zhFontAdvanced } from './font-advanced';
import { zhLigatures } from './ligatures';
import { zhNavigation } from './navigation';
import { zhLineNumbers } from './line-numbers';
import { zhDropCap } from './drop-cap';
import type { LocaleStrings } from './en';
import { zhCNApp } from './zh-CN-app';
import { zhTableProperties } from './table-properties';
import { zhCNRibbon } from './zh-CN-ribbon';

export const zhCN: LocaleStrings = {
	...zhCNRibbon,
	...zhCNApp,
	...zhTableProperties,
	...zhLineNumbers,
	...zhDropCap,
	...zhNavigation,
	...zhColumns,
	...zhMultilevelList,
	...zhFontAdvanced,
	...zhLigatures,
};
