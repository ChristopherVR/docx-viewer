import type { Block, Note, Paragraph, TextRun } from '@christophervr/docx-core';
import { openBlocksEditor, type InlineEditorOptions } from './header-footer-editor';

export interface NoteEditingOptions {
	editor?: InlineEditorOptions;
	note(id: string): Note | undefined;
	change(id: string, blocks: Block[]): void;
	editable(): boolean;
}

/** Splits off the note's automatic number mark so it can't be edited away. */
function withoutMark(blocks: Block[]): { editable: Block[]; mark?: TextRun } {
	const [first, ...rest] = blocks;
	if (first?.type !== 'paragraph' || !first.runs[0]?.noteMark) return { editable: blocks };
	const [mark, ...runs] = first.runs;
	return {
		editable: [{ ...first, runs: runs.length ? runs : [{ text: '' }] }, ...rest],
		mark,
	};
}

function withMark(blocks: Block[], mark: TextRun | undefined): Block[] {
	const [first, ...rest] = blocks;
	if (!mark || first?.type !== 'paragraph') return blocks;
	const paragraph: Paragraph = { ...first, runs: [mark, ...first.runs] };
	return [paragraph, ...rest];
}

/** Double-click a footnote or endnote to edit its text in place. */
export function attachNoteEditing(root: HTMLElement, options: NoteEditingOptions): void {
	for (const item of root.querySelectorAll<HTMLElement>('li[data-docx-note-id]')) {
		const id = item.dataset.docxNoteId!;
		const body = item.querySelector<HTMLElement>('.dve-note-body');
		if (!body) continue;
		item.title = 'Double-click to edit';
		item.addEventListener('dblclick', () => {
			if (!options.editable() || item.classList.contains('dve-header-footer-editing')) return;
			const note = options.note(id);
			if (!note) return;
			const { editable, mark } = withoutMark(note.blocks);
			openBlocksEditor(
				item,
				body,
				editable,
				(blocks) => options.change(id, withMark(blocks, mark)),
				options.editor,
			);
		});
	}
}
