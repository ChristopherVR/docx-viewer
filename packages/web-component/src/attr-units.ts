import {
	roundSignedTwips,
	roundTwips,
	type Paragraph,
	type SignedTwips,
	type Twips,
} from 'docx-core';

// ProseMirror node attributes are untyped, so measurements cross into the branded model here:
// values are rounded to whole twips, and anything that is not a finite number (or is negative
// for an unsigned field) is dropped rather than written into the model.
/** A non-negative twip attribute as `Twips`, or undefined when absent or not a valid length. */
export function twipsAttr(value: unknown): Twips | undefined {
	return typeof value === 'number' && Number.isFinite(value) && value >= 0
		? roundTwips(value)
		: undefined;
}
/** A possibly negative twip attribute as `SignedTwips`, or undefined when absent or not finite. */
export function signedTwipsAttr(value: unknown): SignedTwips | undefined {
	return typeof value === 'number' && Number.isFinite(value) ? roundSignedTwips(value) : undefined;
}

type ParagraphTwips = Pick<
	Paragraph,
	| 'spacingBeforeTwips'
	| 'spacingAfterTwips'
	| 'lineSpacingTwips'
	| 'indentLeftTwips'
	| 'indentRightTwips'
	| 'indentStartTwips'
	| 'indentEndTwips'
	| 'firstLineTwips'
	| 'hangingTwips'
>;
/** The paragraph's direct spacing and indent attributes as model fields; absent ones are omitted. */
export function paragraphTwipsFromAttrs(attrs: Record<string, unknown>): ParagraphTwips {
	const result: ParagraphTwips = {};
	const spacingBefore = twipsAttr(attrs.spacingBeforeTwips);
	const spacingAfter = twipsAttr(attrs.spacingAfterTwips);
	const lineSpacing = signedTwipsAttr(attrs.lineSpacingTwips);
	const left = signedTwipsAttr(attrs.indentLeftTwips);
	const right = signedTwipsAttr(attrs.indentRightTwips);
	const start = signedTwipsAttr(attrs.indentStartTwips);
	const end = signedTwipsAttr(attrs.indentEndTwips);
	const firstLine = twipsAttr(attrs.firstLineTwips);
	const hanging = twipsAttr(attrs.hangingTwips);
	if (spacingBefore !== undefined) result.spacingBeforeTwips = spacingBefore;
	if (spacingAfter !== undefined) result.spacingAfterTwips = spacingAfter;
	if (lineSpacing !== undefined) result.lineSpacingTwips = lineSpacing;
	if (left !== undefined) result.indentLeftTwips = left;
	if (right !== undefined) result.indentRightTwips = right;
	if (start !== undefined) result.indentStartTwips = start;
	if (end !== undefined) result.indentEndTwips = end;
	if (firstLine !== undefined) result.firstLineTwips = firstLine;
	if (hanging !== undefined) result.hangingTwips = hanging;
	return result;
}
