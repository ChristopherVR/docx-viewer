/**
 * Pure helpers for Word's paragraph "keep together" rules. Kept separate
 * from page-flow.ts (which owns page/column bookkeeping) so the rules
 * themselves are trivial to unit test in isolation.
 */

/**
 * Word's default widow/orphan control: never leave a single line of a
 * paragraph stranded alone at the bottom of a page/column (an "orphan") or
 * alone at the top of the next one (a "widow"). Returns how many of the
 * paragraph's `totalLines` should actually be placed on the current
 * page/column, given that `fittingLines` would otherwise fit by height alone.
 */
export function adjustForWidowOrphan(
	totalLines: number,
	fittingLines: number,
	widowControl: boolean,
): number {
	if (!widowControl) return fittingLines;
	if (fittingLines <= 0 || fittingLines >= totalLines) return fittingLines;
	const remaining = totalLines - fittingLines;
	if (fittingLines === 1) return 0; // orphan: the first line would be stranded alone
	if (remaining === 1) return fittingLines - 1; // widow: the last line would start alone
	return fittingLines;
}

/** `w:contextualSpacing`: suppress spacing between adjacent same-style paragraphs. */
export function suppressesSpacing(
	previous: { styleId?: string; contextualSpacing?: boolean } | undefined,
	current: { styleId?: string; contextualSpacing?: boolean },
): boolean {
	if (!previous || !previous.styleId || previous.styleId !== current.styleId) return false;
	return Boolean(previous.contextualSpacing) || Boolean(current.contextualSpacing);
}

/** Word's document default for widow/orphan control is on unless explicitly turned off. */
export function widowControlEnabled(value: boolean | undefined): boolean {
	return value !== false;
}
