/** WordprocessingML ST_HighlightColor tokens, including `none` for clearing. */
export const WORD_HIGHLIGHT_TOKENS = [
	'black',
	'blue',
	'cyan',
	'green',
	'magenta',
	'red',
	'yellow',
	'white',
	'darkBlue',
	'darkCyan',
	'darkGreen',
	'darkMagenta',
	'darkRed',
	'darkYellow',
	'darkGray',
	'lightGray',
	'none',
] as const;

export type WordHighlightToken = (typeof WORD_HIGHLIGHT_TOKENS)[number];

export function isWordHighlightToken(value: string): value is WordHighlightToken {
	return (WORD_HIGHLIGHT_TOKENS as readonly string[]).includes(value);
}
