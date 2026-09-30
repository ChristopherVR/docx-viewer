import { esColumns } from './columns';
import { esMultilevelList } from './multilevel-list';
import { esFontAdvanced } from './font-advanced';
import { esLigatures } from './ligatures';
import { esNavigation } from './navigation';
import { esDropCap } from './drop-cap';
import { esLineNumbers } from './line-numbers';
import type { LocaleStrings } from './en';
import { esApp } from './es-app';
import { esTableProperties } from './table-properties';
import { esRibbon } from './es-ribbon';

export const es: LocaleStrings = {
	...esRibbon,
	...esApp,
	...esTableProperties,
	...esLineNumbers,
	...esDropCap,
	...esNavigation,
	...esColumns,
	...esMultilevelList,
	...esFontAdvanced,
	...esLigatures,
};
