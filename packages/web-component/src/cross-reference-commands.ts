import { headingLevel, type DocumentModel, type Paragraph } from 'docx-core';
import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { CAPTION_LABELS, seqLabelOf } from './caption-commands';
import { schema } from './schema';

export type ReferenceKind = 'heading' | 'bookmark' | 'figure' | 'table' | 'equation';
export type ReferenceContent = 'text' | 'page' | 'label';

export interface ReferenceTarget {
	/** What the list shows. */
	label: string;
	/** Position of the target paragraph. */
	pos: number;
	/** The paragraph's id, for page lookup. */
	id: string;
	/** The whole paragraph text. */
	text: string;
	/** For a caption: "Figure 3". */
	numbered?: string;
	/** An existing bookmark on the paragraph that the field can point at. */
	bookmark?: string;
}

const CAPTION_OF: Partial<Record<ReferenceKind, string>> = {
	figure: CAPTION_LABELS[0],
	table: CAPTION_LABELS[1],
	equation: CAPTION_LABELS[2],
};

const clip = (text: string) => (text.length > 70 ? `${text.slice(0, 67)}…` : text);

/** Every paragraph of `kind`, in document order: headings, bookmarked paragraphs or captions. */
export function referenceTargets(
	doc: ProseMirrorNode,
	model: DocumentModel,
	kind: ReferenceKind,
): ReferenceTarget[] {
	const targets: ReferenceTarget[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		const text = node.textContent.trim();
		const bookmarks = ((node.attrs.bookmarks as string[]) ?? []).filter(Boolean);
		const base = { pos, id: String(node.attrs.id), text };
		if (kind === 'heading') {
			const level = headingLevel(
				{ style: node.attrs.style || undefined } as Paragraph,
				model.paragraphStyles,
			);
			if (level && text)
				targets.push({
					...base,
					label: `${'  '.repeat(level - 1)}${clip(text)}`,
					...(bookmarks[0] ? { bookmark: bookmarks[0] } : {}),
				});
		} else if (kind === 'bookmark') {
			for (const name of bookmarks.filter((item) => !item.startsWith('_')))
				targets.push({ ...base, label: name, bookmark: name });
		} else {
			const wanted = CAPTION_OF[kind]!;
			let number: string | undefined;
			node.forEach((child) => {
				if (seqLabelOf(child)?.toLowerCase() === wanted.toLowerCase()) number = child.text ?? '';
			});
			if (number !== undefined) {
				const numbered = `${text.slice(0, text.indexOf(number))}${number}`.trim();
				targets.push({
					...base,
					label: clip(text),
					numbered,
					...(bookmarks[0] ? { bookmark: bookmarks[0] } : {}),
				});
			}
		}
		return false;
	});
	return targets;
}

let counter = 0;

/** A new hidden `_Ref` bookmark name, the kind Word makes for cross-references. */
export function newReferenceBookmark(existing: readonly string[]): string {
	let name: string;
	do name = `_Ref${(Date.now() % 1_000_000_000) + counter++}`;
	while (existing.includes(name));
	return name;
}

/**
 * Insert > Cross-reference: puts a `REF` (text or caption label and number) or `PAGEREF` field
 * at the selection, pointing at `target`. A paragraph without a bookmark gets a hidden `_Ref`
 * one in the same undoable step. `pageOf` supplies a page number for `page`. The result text is
 * computed now; it does not update when the target changes until the field is edited.
 */
export function insertCrossReference(
	view: EditorView,
	target: ReferenceTarget,
	content: ReferenceContent,
	pageOf: (id: string) => string | undefined,
): boolean {
	if (!view.editable) return false;
	// The target may be stale if the document changed while its dialog was open.
	if (target.pos < 0 || target.pos >= view.state.doc.content.size) return false;
	const node = view.state.doc.nodeAt(target.pos);
	if (node?.type.name !== 'paragraph') return false;
	const names = ((node.attrs.bookmarks as string[]) ?? []).filter(Boolean);
	const name = target.bookmark ?? newReferenceBookmark(names);
	let tr = view.state.tr;
	if (!names.includes(name)) tr = tr.setNodeAttribute(target.pos, 'bookmarks', [...names, name]);
	const result =
		content === 'page'
			? (pageOf(target.id) ?? '1')
			: content === 'label'
				? (target.numbered ?? target.text)
				: target.text;
	const instr = content === 'page' ? ` PAGEREF ${name} \\h ` : ` REF ${name} \\h `;
	const field = schema.text(result || name, [
		...(view.state.storedMarks ?? []),
		schema.marks.field!.create({ instr, simple: true }),
	]);
	tr = tr.replaceSelectionWith(field, false);
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}
