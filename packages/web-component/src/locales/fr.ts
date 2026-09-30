import { frColumns } from './columns';
import { frMultilevelList } from './multilevel-list';
import { frFontAdvanced } from './font-advanced';
import { frLigatures } from './ligatures';
import { frNavigation } from './navigation';
import { frDropCap } from './drop-cap';
import { frLineNumbers } from './line-numbers';
import type { LocaleStrings } from './en';
import { frApp } from './fr-app';
import { frTableProperties } from './table-properties';
import { frRibbon } from './fr-ribbon';

export const fr: LocaleStrings = {
	...frRibbon,
	...frApp,
	...frTableProperties,
	...frLineNumbers,
	...frDropCap,
	...frNavigation,
	...frColumns,
	...frMultilevelList,
	...frFontAdvanced,
	...frLigatures,
};
