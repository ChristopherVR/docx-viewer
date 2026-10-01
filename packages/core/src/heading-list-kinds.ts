// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { StNumberFormat } from './generated/wml-simple-types.js';
import type { NumberingLevelDefinition } from './numbering-model.js';
import { signedTwips, twips } from './units.js';

/** Heading-linked outline lists from Word's Multilevel List gallery. */
export type HeadingListKind = 'headings' | 'article' | 'roman';

type Spec = readonly [format: StNumberFormat, text: string, legal?: boolean];

const SPECS: Record<HeadingListKind, readonly Spec[]> = {
	// 1 Heading 1 / 1.1 Heading 2 / 1.1.1 Heading 3 ...
	headings: Array.from({ length: 9 }, (_, level): Spec => [
		'decimal',
		Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join('.'),
	]),
	// Article I. / Section 1.1 (Arabic) / (a) / (i)
	article: [
		['upperRoman', 'Article %1.'],
		['decimal', 'Section %1.%2', true],
		['lowerLetter', '(%3)'],
		['lowerRoman', '(%4)'],
		['upperLetter', '(%5)'],
		['decimal', '(%6)'],
		['lowerLetter', '(%7)'],
		['lowerRoman', '(%8)'],
		['upperLetter', '(%9)'],
	],
	// I. / A. / 1. / a) / (1) / (a) / (i) ...
	roman: [
		['upperRoman', '%1.'],
		['upperLetter', '%2.'],
		['decimal', '%3.'],
		['lowerLetter', '%4)'],
		['decimal', '(%5)'],
		['lowerLetter', '(%6)'],
		['lowerRoman', '(%7)'],
		['lowerLetter', '(%8)'],
		['lowerRoman', '(%9)'],
	],
};

export const isHeadingListKind = (value: string): value is HeadingListKind => value in SPECS;

/**
 * Nine levels for a heading-linked list. Level `n` names the built-in `Heading n+1` style; callers
 * drop `paragraphStyleId` for styles the document does not have.
 */
export function headingListLevels(kind: HeadingListKind): NumberingLevelDefinition[] {
	return SPECS[kind].map(([numFmt, lvlText, legal], level) => {
		const indent = 432 + 144 * level;
		return {
			level,
			start: 1,
			numFmt,
			lvlText,
			...(legal ? { isLgl: true } : {}),
			paragraphStyleId: `Heading${level + 1}`,
			indentLeftTwips: signedTwips(indent),
			hangingTwips: twips(indent),
			suffix: 'tab' as const,
		};
	});
}
