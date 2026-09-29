import {
	signedTwips,
	twips,
	type DocumentModel,
	type SectionProperties,
} from '@christophervr/docx-core';
import { pageSizeOf } from './page-size';
import { withSection } from './section-commands';

const TWIPS_PER_INCH = 1440;

/** What Word's Page Setup dialog edits for one section, in inches. */
export interface PageSetupValues {
	topIn: number;
	bottomIn: number;
	leftIn: number;
	rightIn: number;
	gutterIn: number;
	headerIn: number;
	footerIn: number;
	orientation: 'portrait' | 'landscape';
	/** Portrait width and height; landscape swaps them when applied. */
	widthIn: number;
	heightIn: number;
}

export type PageSetupProblem = 'range' | 'width' | 'height';

const MAX_PAGE_INCHES = 22;
const MIN_TEXT_INCHES = 0.5;

const inches = (value: number) => Math.round((value / TWIPS_PER_INCH) * 1000) / 1000;

/** The section's current settings, with the paper always in portrait terms. */
export function readPageSetup(section: SectionProperties): PageSetupValues {
	const short = Math.min(section.pageWidthTwips, section.pageHeightTwips);
	const long = Math.max(section.pageWidthTwips, section.pageHeightTwips);
	return {
		topIn: inches(section.marginTopTwips),
		bottomIn: inches(section.marginBottomTwips),
		leftIn: inches(section.marginLeftTwips),
		rightIn: inches(section.marginRightTwips),
		gutterIn: inches(section.gutterTwips ?? 0),
		headerIn: inches(section.headerDistanceTwips ?? 720),
		footerIn: inches(section.footerDistanceTwips ?? 720),
		orientation: section.orientation,
		widthIn: inches(short),
		heightIn: inches(long),
	};
}

/**
 * Checks the values as Word does: every measure in range, and margins that leave room for text
 * (at least half an inch each way). Returns the first problem, or null when they can be applied.
 */
export function validatePageSetup(values: PageSetupValues): PageSetupProblem | null {
	const margins = [values.topIn, values.bottomIn, values.leftIn, values.rightIn, values.gutterIn];
	const measures = [...margins, values.headerIn, values.footerIn];
	if (measures.some((value) => !Number.isFinite(value) || value < 0 || value > MAX_PAGE_INCHES))
		return 'range';
	if (![values.widthIn, values.heightIn].every((v) => v > 0 && v <= MAX_PAGE_INCHES))
		return 'range';
	const short = Math.min(values.widthIn, values.heightIn);
	const long = Math.max(values.widthIn, values.heightIn);
	const [width, height] = values.orientation === 'landscape' ? [long, short] : [short, long];
	if (width - values.leftIn - values.rightIn - values.gutterIn < MIN_TEXT_INCHES) return 'width';
	if (height - values.topIn - values.bottomIn < MIN_TEXT_INCHES) return 'height';
	return null;
}

const toTwips = (value: number) => Math.round(value * TWIPS_PER_INCH);

/** Applies validated values to section `index`; invalid values leave the model as it was. */
export function applyPageSetup(
	model: DocumentModel,
	index: number,
	values: PageSetupValues,
): DocumentModel {
	if (validatePageSetup(values)) return model;
	const short = Math.min(values.widthIn, values.heightIn);
	const long = Math.max(values.widthIn, values.heightIn);
	const landscape = values.orientation === 'landscape';
	return withSection(model, index, (section) => ({
		...section,
		orientation: values.orientation,
		pageWidthTwips: twips(toTwips(landscape ? long : short)),
		pageHeightTwips: twips(toTwips(landscape ? short : long)),
		marginTopTwips: signedTwips(toTwips(values.topIn)),
		marginBottomTwips: signedTwips(toTwips(values.bottomIn)),
		marginLeftTwips: twips(toTwips(values.leftIn)),
		marginRightTwips: twips(toTwips(values.rightIn)),
		gutterTwips: twips(toTwips(values.gutterIn)),
		headerDistanceTwips: twips(toTwips(values.headerIn)),
		footerDistanceTwips: twips(toTwips(values.footerIn)),
	}));
}

/** The paper-size preset a set of values matches, or null for a custom size. */
export function presetOf(values: PageSetupValues): string | null {
	return pageSizeOf({
		pageWidthTwips: twips(toTwips(values.widthIn)),
		pageHeightTwips: twips(toTwips(values.heightIn)),
	} as SectionProperties);
}
