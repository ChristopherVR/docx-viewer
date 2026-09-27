// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model.js';

/** Direct run properties supported by the editor for style/default resolution, in native Word units. */
export type RunFormatting = Pick<
	TextRun,
	| 'bold'
	| 'italic'
	| 'underline'
	| 'underlineStyle'
	| 'underlineColor'
	| 'strike'
	| 'doubleStrike'
	| 'caps'
	| 'smallCaps'
	| 'vanish'
	| 'highlight'
	| 'verticalAlign'
	| 'fontSize'
	| 'fontFamily'
	| 'fontTheme'
	| 'color'
	| 'colorTheme'
	| 'characterSpacingTwips'
	| 'shadingFill'
	| 'shadingThemeFill'
>;
/** A `styles.xml` style carrying run formatting: a character style, or a paragraph style's `w:rPr`. */
export interface CharacterStyleDefinition {
	id: string;
	type: 'character' | 'paragraph';
	name?: string;
	basedOn?: string;
	/** `w:link`; the paired style of the opposite kind sharing this style's look. */
	linkedStyle?: string;
	isDefault?: boolean;
	formatting: RunFormatting;
}
/** Parsed run-level style data (docDefaults `rPrDefault` plus character/paragraph style `rPr`). Read-only catalog. */
export interface RunStyleCatalog {
	docDefaults: RunFormatting;
	styles: Record<string, CharacterStyleDefinition>;
	warnings: string[];
}
