import type { WordUnderlineStyle } from './theme-model.js';

/** WordprocessingML `ST_Underline` tokens, including `none` for clearing. */
export const WORD_UNDERLINE_TOKENS = [
	'single',
	'words',
	'double',
	'thick',
	'dotted',
	'dottedHeavy',
	'dash',
	'dashedHeavy',
	'dashLong',
	'dashLongHeavy',
	'dotDash',
	'dashDotHeavy',
	'dotDotDash',
	'dashDotDotHeavy',
	'wave',
	'wavyHeavy',
	'wavyDouble',
	'none',
] as const satisfies readonly WordUnderlineStyle[];

export function isWordUnderlineStyle(value: string): value is WordUnderlineStyle {
	return (WORD_UNDERLINE_TOKENS as readonly string[]).includes(value);
}
