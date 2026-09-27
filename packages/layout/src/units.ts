// Framework-neutral pagination engine; conversions match the CSS-pixel-at-96dpi
// convention already used by @christophervr/docx-core's DocumentModel.page.
/** Twentieths of a point (Word's native unit for spacing/indents). */
export function twipsToPx(twips: number): number {
	return (twips * 96) / 1440;
}
export function pxToTwips(px: number): number {
	return (px * 1440) / 96;
}
/** Points (Word's native unit for font sizes). */
export function ptToPx(pt: number): number {
	return (pt * 96) / 72;
}
/** Default tab stop interval: every half inch when no custom `w:tabs` are modeled. */
export const DEFAULT_TAB_STOP_PX = twipsToPx(720);
/** Word's default body font when a run specifies neither family nor size. */
export const DEFAULT_FONT_FAMILY = 'Calibri';
export const DEFAULT_FONT_SIZE_PT = 11;
/** Single line-spacing is expressed by Word as 240 twentieths-of-a-line. */
export const SINGLE_LINE_SPACING_UNITS = 240;
