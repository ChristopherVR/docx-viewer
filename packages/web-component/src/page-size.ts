import { twips, type DocumentModel, type SectionProperties } from 'docx-core';
import { withSection } from './section-commands';

/** Word's Layout > Size presets, portrait width by height in twips (1/1440 inch). */
export const PAGE_SIZES: Record<string, readonly [width: number, height: number]> = {
	letter: [12240, 15840],
	legal: [12240, 20160],
	tabloid: [15840, 24480],
	executive: [10440, 15120],
	a3: [16838, 23811],
	a4: [11906, 16838],
	a5: [8391, 11906],
	b5: [10319, 14570],
};

/** English labels for the size list; paper names are proper nouns and stay as they are. */
export const PAGE_SIZE_OPTIONS: Array<[string, string]> = [
	['letter', 'Letter'],
	['legal', 'Legal'],
	['tabloid', 'Tabloid'],
	['executive', 'Executive'],
	['a3', 'A3'],
	['a4', 'A4'],
	['a5', 'A5'],
	['b5', 'B5 (JIS)'],
];

/** Sets a section's paper size, keeping its orientation (landscape swaps width and height). */
export function setPageSize(model: DocumentModel, index: number, preset: string): DocumentModel {
	const size = PAGE_SIZES[preset];
	if (!size) return model;
	return withSection(model, index, (section) => {
		const [short, long] = size;
		const landscape = section.orientation === 'landscape';
		return {
			...section,
			pageWidthTwips: twips(landscape ? long : short),
			pageHeightTwips: twips(landscape ? short : long),
		};
	});
}

/** The preset a section's paper matches in either orientation, or null for a custom size. */
export function pageSizeOf(section: SectionProperties): string | null {
	const a = Math.min(section.pageWidthTwips, section.pageHeightTwips);
	const b = Math.max(section.pageWidthTwips, section.pageHeightTwips);
	for (const [key, [short, long]] of Object.entries(PAGE_SIZES))
		if (Math.abs(short - a) <= 2 && Math.abs(long - b) <= 2) return key;
	return null;
}
