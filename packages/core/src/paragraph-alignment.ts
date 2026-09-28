// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The single `w:pPr/w:jc` parser shared by direct paragraphs, styles and docDefaults.
import type { StJc } from './generated/wml-simple-types.js';
import { isStJc } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';
import { onOffElement } from './simple-types.js';
import { first, getW, type XmlElement } from './xml.js';

export type ParagraphAlign = 'left' | 'center' | 'right' | 'justify';

/** Maps a full `ST_Jc` value to the renderer's alignment; `start`/`end` follow the paragraph direction. */
export function alignFromJustification(jc: StJc, rtl: boolean): ParagraphAlign | undefined {
	switch (jc) {
		case 'left':
		case 'center':
		case 'right':
			return jc;
		case 'start':
			return rtl ? 'right' : 'left';
		case 'end':
			return rtl ? 'left' : 'right';
		case 'both':
		case 'distribute':
		case 'thaiDistribute':
		case 'mediumKashida':
		case 'highKashida':
		case 'lowKashida':
			return 'justify';
		default:
			return undefined; // numTab aligns to the list tab and has no plain equivalent.
	}
}

export interface ParsedJustification {
	/** Renderer alignment derived from `justification` and the local `w:bidi`. */
	align?: ParagraphAlign;
	/** The exact `ST_Jc` value read from the file, kept so it round-trips unchanged. */
	justification?: StJc;
}

/** Parses `w:jc` in a `w:pPr`; invalid values are dropped with a parse warning. */
export function parseJustification(pPr: XmlElement | undefined): ParsedJustification {
	const justification = enumValue(isStJc, getW(first(pPr, 'jc'), 'val'), 'w:jc');
	if (!justification) return {};
	const align = alignFromJustification(justification, onOffElement(first(pPr, 'bidi')) === true);
	return { justification, ...(align ? { align } : {}) };
}
