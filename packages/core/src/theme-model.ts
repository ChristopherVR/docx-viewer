// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
/** Word theme color scheme slots (`a:clrScheme` children), in schema order. */
export type ThemeColorSlot =
	| 'dk1'
	| 'lt1'
	| 'dk2'
	| 'lt2'
	| 'accent1'
	| 'accent2'
	| 'accent3'
	| 'accent4'
	| 'accent5'
	| 'accent6'
	| 'hlink'
	| 'folHlink';
/** WordprocessingML `ST_ThemeColor` tokens referenced from `w:themeColor`/`w:themeFill`. */
export type ThemeColorToken =
	| 'dark1'
	| 'light1'
	| 'dark2'
	| 'light2'
	| 'accent1'
	| 'accent2'
	| 'accent3'
	| 'accent4'
	| 'accent5'
	| 'accent6'
	| 'hyperlink'
	| 'followedHyperlink'
	| 'background1'
	| 'text1'
	| 'background2'
	| 'text2';
export interface ThemeFontSet {
	latin?: string;
	eastAsia?: string;
	complexScript?: string;
}
/** Parsed word/theme/theme1.xml plus word/settings.xml clrSchemeMapping; never flattened onto runs. */
export interface ThemeCatalog {
	colors: Partial<Record<ThemeColorSlot, string>>;
	colorMapping: Partial<Record<'bg1' | 'tx1' | 'bg2' | 'tx2', ThemeColorSlot>>;
	fonts: { major: ThemeFontSet; minor: ThemeFontSet };
}
/** A direct theme color reference kept alongside (not instead of) any resolved RGB fallback. */
export interface ThemeColorReference {
	token: ThemeColorToken;
	/** Fraction in [0,1], parsed from the two-digit hex byte Word stores. */
	tint?: number;
	shade?: number;
}
export type ThemeFontScript = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';
export type ThemeFontRole = 'major' | 'minor';
/** Word `ST_Underline` styles beyond the plain single-line toggle. */
export type WordUnderlineStyle =
	| 'single'
	| 'words'
	| 'double'
	| 'thick'
	| 'dotted'
	| 'dottedHeavy'
	| 'dash'
	| 'dashedHeavy'
	| 'dashLong'
	| 'dashLongHeavy'
	| 'dotDash'
	| 'dashDotHeavy'
	| 'dotDotDash'
	| 'dashDotDotHeavy'
	| 'wave'
	| 'wavyHeavy'
	| 'wavyDouble'
	| 'none';
