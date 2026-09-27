// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Tags field result runs with their field code so renderers can recalculate display-only fields
// such as PAGE and NUMPAGES. The markers and field code are modeled as their own runs.
import type { TextRun } from './model.js';
import { getW, isElement, named, textContent, type XmlElement } from './xml.js';

/** The field code's name (first word, upper case), e.g. `PAGE` for ` PAGE \* MERGEFORMAT `. */
export function fieldName(instr: string): string {
	return instr.trim().split(/\s+/)[0]?.toUpperCase() ?? '';
}

/**
 * Follows `w:fldChar` begin/separate/end markers across a paragraph's runs. Runs between
 * `separate` and `end` form the field's cached result and receive `field: { instr }`.
 */
export function createFieldTracker() {
	let depth = 0;
	let instr = '';
	let inResult = false;
	return (element: XmlElement, run: TextRun): TextRun => {
		let tagged = false;
		for (const child of Array.from(element.childNodes)) {
			if (!isElement(child)) continue;
			if (named(child, 'fldChar')) {
				const type = getW(child, 'fldCharType');
				if (type === 'begin') {
					depth += 1;
					if (depth === 1) {
						instr = '';
						inResult = false;
					}
				} else if (type === 'separate' && depth === 1) inResult = true;
				else if (type === 'end' && depth > 0) {
					depth -= 1;
					if (depth === 0) inResult = false;
				}
			} else if (named(child, 'instrText') && depth === 1 && !inResult) instr += textContent(child);
			else if (named(child, 't') && inResult && depth === 1) tagged = true;
		}
		if (tagged && instr.trim()) run.field = { instr: instr.trim() };
		return run;
	};
}
