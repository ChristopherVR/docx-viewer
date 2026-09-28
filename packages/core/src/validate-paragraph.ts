// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Pre-save validation of paragraphs and runs against the ECMA-376 simple types.
import {
	isStBrType,
	isStFldCharType,
	isStHighlightColor,
	isStJc,
	isStLineSpacingRule,
	isStTabJc,
	isStTabTlc,
	isStUnderline,
} from './generated/wml-simple-types.js';
import type { Paragraph, Revision, TabStop, TextRun } from './model.js';
import type { ThemeColorReference } from './theme-model.js';
import { Checker, show } from './validate-issues.js';

const isFraction = (value: unknown): boolean =>
	typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 1;

/** Validates a tint/shade-carrying theme reference (`w:color/@w:themeColor` and friends). */
export function validateThemeReference(
	c: Checker,
	field: string,
	value: ThemeColorReference | undefined,
): void {
	if (value === undefined) return;
	const at = c.at(`.${field}`);
	at.theme('token', value.token);
	at.check('tint', value.tint, isFraction, 'must be a fraction between 0 and 1');
	at.check('shade', value.shade, isFraction, 'must be a fraction between 0 and 1');
}

/** Validates the tracked-change metadata written as `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`. */
export function validateRevision(c: Checker, field: string, revision: Revision | undefined): void {
	if (revision === undefined) return;
	c.at(`.${field}`).dateTime('date', revision.date);
}

/** `w:sz` is `ST_HpsMeasure`: an unsigned integer of half-points, so the point size must be finite and non-negative. */
function validateFontSize(c: Checker, value: unknown): void {
	c.check(
		'fontSize',
		value,
		(v) => typeof v === 'number' && Number.isFinite(v) && Math.round(v * 2) >= 0,
		'must be a finite, non-negative point size that converts to whole half-points (ST_HpsMeasure)',
	);
}

export function validateRun(c: Checker, run: TextRun): void {
	validateFontSize(c, run.fontSize);
	c.hex('color', run.color);
	c.hex('shadingFill', run.shadingFill);
	c.hex('underlineColor', run.underlineColor);
	validateThemeReference(c, 'colorTheme', run.colorTheme);
	validateThemeReference(c, 'shadingThemeFill', run.shadingThemeFill);
	c.enum('underlineStyle', run.underlineStyle, isStUnderline, 'ST_Underline');
	if (run.highlight !== undefined && !isStHighlightColor(run.highlight))
		c.fail(
			'highlight',
			run.highlight,
			`Unsupported Word highlight token: ${String(run.highlight)}`,
		);
	c.enum('break', run.break, isStBrType, 'ST_BrType');
	c.enum('fieldChar', run.fieldChar, isStFldCharType, 'ST_FldCharType');
	c.check(
		'characterSpacingTwips',
		run.characterSpacingTwips,
		(v) => typeof v === 'number' && Number.isFinite(v),
		'must be a finite number of twips (ST_SignedTwipsMeasure)',
	);
	c.language('language', run.language);
	c.language('eastAsiaLanguage', run.eastAsiaLanguage);
	c.language('bidiLanguage', run.bidiLanguage);
	c.check(
		'verticalAlign',
		run.verticalAlign,
		(v) => v === 'superscript' || v === 'subscript',
		'must be "superscript" or "subscript"',
	);
	validateRevision(c, 'revision', run.revision);
	if (run.image) {
		const image = c.at('.image');
		for (const key of ['widthPx', 'heightPx'] as const)
			image.check(
				key,
				run.image[key],
				(v) => typeof v === 'number' && Number.isFinite(v) && v >= 0,
				'must be a finite, non-negative pixel size',
			);
	}
}

function validateTabStop(c: Checker, stop: TabStop): void {
	c.signed('posTwips', stop.posTwips, 'CT_TabStop/@pos');
	c.enum('align', stop.align, isStTabJc, 'ST_TabJc');
	c.enum('leader', stop.leader, isStTabTlc, 'ST_TabTlc');
}

export function validateParagraph(c: Checker, paragraph: Paragraph): void {
	c.check(
		'align',
		paragraph.align,
		(v) => v === 'left' || v === 'center' || v === 'right' || v === 'justify',
		'must be "left", "center", "right" or "justify"',
	);
	c.enum('justification', paragraph.justification, isStJc, 'ST_Jc');
	c.check(
		'direction',
		paragraph.direction,
		(v) => v === 'ltr' || v === 'rtl',
		'must be "ltr" or "rtl"',
	);
	for (const key of [
		'spacingBeforeTwips',
		'spacingAfterTwips',
		'firstLineTwips',
		'hangingTwips',
	] as const)
		c.unsigned(key, paragraph[key], 'ST_TwipsMeasure');
	for (const key of [
		'lineSpacingTwips',
		'indentLeftTwips',
		'indentRightTwips',
		'indentStartTwips',
		'indentEndTwips',
	] as const)
		c.signed(key, paragraph[key], 'ST_SignedTwipsMeasure');
	c.enum('lineSpacingRule', paragraph.lineSpacingRule, isStLineSpacingRule, 'ST_LineSpacingRule');
	if (paragraph.numbering) {
		const numbering = c.at('.numbering');
		numbering.unsigned('numId', paragraph.numbering.numId, 'ST_DecimalNumber');
		numbering.unsigned('level', paragraph.numbering.level, 'ST_DecimalNumber');
	}
	paragraph.tabStops?.forEach((stop, index) => validateTabStop(c.at(`.tabStops[${index}]`), stop));
	validateRevision(c, 'markRevision', paragraph.markRevision);
	paragraph.runs.forEach((run, index) => validateRun(c.at(`.runs[${index}]`), run));
}
