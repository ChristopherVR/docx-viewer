/**
 * A UAX#14-lite tokenizer: enough of the Unicode line-breaking algorithm to
 * place break opportunities at spaces, hyphens, and between CJK characters,
 * without implementing the full class table.
 *
 * The `isEastAsianChar` codepoint ranges are ported, with this provenance
 * note, from the sibling PowerPoint viewer's East Asian line-breaking helper
 * at `../pptx-viewer-new/packages/angular/src/internal/shared-src/render/text-east-asian-breaks.ts`
 * (`isEastAsianChar`). That module's hanging-punctuation/kinsoku machinery is
 * PowerPoint-specific (`a:pPr/@hangingPunct`, `@eaLnBrk`) and is not ported;
 * only the character classification is reused here.
 */
export function isEastAsianChar(ch: string): boolean {
	const cp = ch.codePointAt(0) ?? 0;
	return (
		(cp >= 0x1100 && cp <= 0x11ff) ||
		(cp >= 0x2e80 && cp <= 0x9fff) ||
		(cp >= 0xa960 && cp <= 0xa97f) ||
		(cp >= 0xac00 && cp <= 0xd7ff) ||
		(cp >= 0xf900 && cp <= 0xfaff) ||
		(cp >= 0xfe30 && cp <= 0xfe4f) ||
		(cp >= 0xff00 && cp <= 0xffef) ||
		(cp >= 0x20000 && cp <= 0x3ffff)
	);
}

export type BreakToken =
	| { kind: 'word'; text: string; runIndex: number; sourceStart: number }
	| { kind: 'space'; text: string; runIndex: number; sourceStart: number }
	| { kind: 'tab'; runIndex: number; sourceStart: number }
	/** An inline picture: an unbreakable box of the run's object size. */
	| { kind: 'object'; runIndex: number; sourceStart: number }
	/** A manual line break (`w:br`/`w:cr`); ends the current line without ending the paragraph. */
	| { kind: 'lineBreak'; runIndex: number; sourceStart: number }
	/** An explicit page or column break carried on a run (see {@link LayoutRun.breakAfter}). */
	| { kind: 'pageBreak' | 'columnBreak'; runIndex: number; sourceStart: number };

/**
 * Splits one run's text into break-opportunity tokens. `sourceStart` is the
 * character offset of each token within the run's own text, so callers can
 * map a line's fragments back to a paragraph/run character range (used for
 * click-to-place-cursor mapping in the Print Layout view).
 */
export function tokenizeRun(text: string, runIndex: number): BreakToken[] {
	const tokens: BreakToken[] = [];
	let word = '';
	let wordStart = 0;
	const flushWord = () => {
		if (word) tokens.push({ kind: 'word', text: word, runIndex, sourceStart: wordStart });
		word = '';
	};
	const chars = [...text];
	let offset = 0;
	for (const ch of chars) {
		if (ch === '\n') {
			flushWord();
			tokens.push({ kind: 'lineBreak', runIndex, sourceStart: offset });
		} else if (ch === '\t') {
			flushWord();
			tokens.push({ kind: 'tab', runIndex, sourceStart: offset });
		} else if (ch === ' ') {
			flushWord();
			tokens.push({ kind: 'space', text: ch, runIndex, sourceStart: offset });
		} else if (isEastAsianChar(ch)) {
			flushWord();
			tokens.push({ kind: 'word', text: ch, runIndex, sourceStart: offset });
		} else {
			if (!word) wordStart = offset;
			word += ch;
			// A plain hyphen is a break opportunity; U+2011 (non-breaking hyphen,
			// `w:noBreakHyphen`) falls through here and is never split on.
			if (ch === '-') flushWord();
		}
		offset += ch.length;
	}
	flushWord();
	return tokens;
}

/** Appends a forced page or column break marker after a run's tokens. */
export function appendBreakMarker(
	tokens: BreakToken[],
	runIndex: number,
	sourceStart: number,
	kind: 'page' | 'column',
): void {
	tokens.push({ kind: kind === 'page' ? 'pageBreak' : 'columnBreak', runIndex, sourceStart });
}
