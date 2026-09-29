// Framework-neutral pagination engine; conversions match the CSS-pixel-at-96dpi
// convention already used by @christophervr/docx-core's DocumentModel.page.
import { roundSignedTwips, twips, type SignedTwips } from '@christophervr/docx-core';

/** Twentieths of a point (Word's native unit for spacing/indents) to CSS pixels. */
export function twipsToPx(value: SignedTwips): number {
	return (value * 96) / 1440;
}
/** Zero twips, the default for absent spacing and indents. */
export const NO_TWIPS = twips(0);
/** CSS pixels to whole twips (rounded half away from zero; may be negative). */
export function pxToTwips(px: number): SignedTwips {
	return roundSignedTwips((px * 1440) / 96);
}
/** Points (Word's native unit for font sizes). */
export function ptToPx(pt: number): number {
	return (pt * 96) / 72;
}
/** Default tab stop interval: every half inch when no custom `w:tabs` are modeled. */
export const DEFAULT_TAB_STOP_PX = twipsToPx(twips(720));
/** Word's default body font when a run specifies neither family nor size. */
export const DEFAULT_FONT_FAMILY = 'Calibri';
export const DEFAULT_FONT_SIZE_PT = 11;
/** Single line-spacing is expressed by Word as 240 twentieths-of-a-line. */
export const SINGLE_LINE_SPACING_UNITS = 240;
