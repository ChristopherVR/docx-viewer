import { closeHistory } from 'prosemirror-history';
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { renumberCaptions, seqLabelOf } from './caption-commands';
import { schema } from './schema';

const REFERENCE = /^\s*(REF|PAGEREF)\s+("[^"]+"|\S+)(.*)$/i;

interface FieldRun {
	from: number;
	to: number;
	mark: Mark;
	text: string;
	marks: readonly Mark[];
}

/**
 * Switches that do not change a REF/PAGEREF result (`\h` hyperlink, `\* MERGEFORMAT`). Any other
 * switch (`\n`, `\r`, `\w`, `\p`, `\f`...) changes what the field shows, so those fields are left
 * alone rather than being given a wrong result.
 */
function plainSwitches(rest: string): boolean {
	return !/\\/.test(rest.replace(/\\h\b/gi, '').replace(/\\\*\s+\S+/g, ''));
}

/** Every body field result as one run: adjacent text nodes carrying the same field mark. */
function fieldRuns(doc: ProseMirrorNode): FieldRun[] {
	const runs: FieldRun[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		let open: FieldRun | undefined;
		node.forEach((child, offset) => {
			const mark = child.isText
				? child.marks.find((item) => item.type === schema.marks.field)
				: undefined;
			const start = pos + 1 + offset;
			if (mark && open?.mark.eq(mark)) {
				open.to = start + child.nodeSize;
				open.text += child.text ?? '';
			} else if (mark) {
				open = {
					from: start,
					to: start + child.nodeSize,
					mark,
					text: child.text ?? '',
					marks: child.marks,
				};
				runs.push(open);
			} else open = undefined;
		});
		return false;
	});
	return runs;
}

/** The paragraph a bookmark sits on, with its id and text. */
function bookmarkTargets(doc: ProseMirrorNode): Map<string, ProseMirrorNode> {
	const targets = new Map<string, ProseMirrorNode>();
	doc.descendants((node) => {
		if (node.type.name !== 'paragraph') return true;
		for (const name of (node.attrs.bookmarks as string[]) ?? [])
			if (name && !targets.has(name)) targets.set(name, node);
		return false;
	});
	return targets;
}

/** "Figure 3": the caption text up to and including its SEQ number, or undefined without one. */
function captionLabel(paragraph: ProseMirrorNode): string | undefined {
	const text = paragraph.textContent;
	let number: string | undefined;
	paragraph.forEach((child) => {
		if (seqLabelOf(child) !== undefined) number = child.text ?? '';
	});
	return number === undefined
		? undefined
		: text.slice(0, text.indexOf(number) + number.length).trim();
}

/**
 * Word's Update Field for the whole body: every `SEQ` caption is renumbered, then `REF` results
 * follow their bookmarked paragraph and `PAGEREF` results take the page from `pageOf`. A REF that
 * pointed at a caption's label and number keeps showing the label; a longer result is treated as
 * the entire caption. Fields with other switches or a missing bookmark are left as they are.
 * One undoable step; returns whether anything changed.
 */
export function updateFields(
	view: EditorView,
	pageOf: (id: string) => string | undefined,
): boolean {
	if (!view.editable) return false;
	const tr = view.state.tr;
	for (const label of new Set(
		fieldRuns(tr.doc)
			.map((run) => seqLabelOf(schema.text('x', [run.mark])))
			.filter((label): label is string => label !== undefined)
			.map((label) => label.toLowerCase()),
	))
		renumberCaptions(tr, label);
	const targets = bookmarkTargets(tr.doc);
	const changes: Array<{ run: FieldRun; text: string }> = [];
	for (const run of fieldRuns(tr.doc)) {
		const match = REFERENCE.exec(String(run.mark.attrs.instr ?? ''));
		if (!match || !plainSwitches(match[3] ?? '')) continue;
		const target = targets.get((match[2] ?? '').replace(/^"|"$/g, ''));
		if (!target) continue;
		let text: string | undefined;
		if (match[1]!.toUpperCase() === 'PAGEREF') text = pageOf(String(target.attrs.id));
		else {
			const label = captionLabel(target);
			const whole = target.textContent.trim();
			text = label !== undefined && run.text.trim().length <= label.length ? label : whole;
		}
		if (text && text !== run.text) changes.push({ run, text });
	}
	// Later fields first, so replacing one never shifts the positions still to be processed.
	for (const { run, text } of changes.reverse())
		tr.replaceWith(run.from, run.to, schema.text(text, run.marks));
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr));
	return true;
}
