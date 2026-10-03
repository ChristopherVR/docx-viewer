import type { TextRun } from 'docx-core';
import { extraRunProperties } from './run-extra-mark';

export const sameJson = (left: unknown, right: unknown): boolean =>
	JSON.stringify(left ?? null) === JSON.stringify(right ?? null);
/** `revision`/`commentIds` are objects/arrays; compare by value rather than by reference. */
function sameRevisionMeta(left: TextRun, right: TextRun): boolean {
	return (
		JSON.stringify(left.revision) === JSON.stringify(right.revision) &&
		JSON.stringify(left.commentIds) === JSON.stringify(right.commentIds) &&
		JSON.stringify(extraRunProperties(left)) === JSON.stringify(extraRunProperties(right)) &&
		JSON.stringify(left.field) === JSON.stringify(right.field) &&
		left.fieldChar === right.fieldChar &&
		left.fieldCode === right.fieldCode
	);
}

/** Whether two run lists render identically (adjacent same-format runs compare merged). */
export function sameRuns(left: TextRun[], right: TextRun[]) {
	const compact = (runs: TextRun[]) =>
		runs.reduce<TextRun[]>((result, run) => {
			// An explicit off (`false`) differs from unset: it cancels a style's value.
			const normalized = { ...run };
			const previous = result.at(-1);
			const fields: (keyof TextRun)[] = [
				'bold',
				'italic',
				'underline',
				'strike',
				'highlight',
				'verticalAlign',
				'language',
				'eastAsiaLanguage',
				'bidiLanguage',
				'rtl',
				'fontFamily',
				'fontSize',
				'color',
				'style',
				'break',
			];
			if (
				previous &&
				!previous.noteReference &&
				!normalized.noteReference &&
				!previous.fieldChar &&
				!normalized.fieldChar &&
				previous.fieldCode === undefined &&
				normalized.fieldCode === undefined &&
				!previous.image &&
				!normalized.image &&
				!previous.equation &&
				!normalized.equation &&
				sameJson(previous.link, normalized.link) &&
				fields.every((field) => previous[field] === normalized[field]) &&
				sameRevisionMeta(previous, normalized)
			)
				previous.text += normalized.text;
			else result.push(normalized);
			return result;
		}, []);
	const a = compact(left);
	const b = compact(right);
	const fields: (keyof TextRun)[] = [
		'bold',
		'italic',
		'underline',
		'strike',
		'highlight',
		'verticalAlign',
		'language',
		'eastAsiaLanguage',
		'bidiLanguage',
		'rtl',
		'fontFamily',
		'fontSize',
		'color',
		'style',
		'break',
	];
	const sameNoteReference = (x?: TextRun['noteReference'], y?: TextRun['noteReference']) =>
		x?.kind === y?.kind && x?.id === y?.id;
	return (
		a.length === b.length &&
		a.every((run, index) => {
			const other = b[index];
			return (
				other !== undefined &&
				run.text === other.text &&
				fields.every((field) => run[field] === other[field]) &&
				sameNoteReference(run.noteReference, other.noteReference) &&
				sameJson(run.link, other.link) &&
				sameJson(run.image, other.image) &&
				sameJson(run.equation, other.equation) &&
				sameRevisionMeta(run, other)
			);
		})
	);
}
