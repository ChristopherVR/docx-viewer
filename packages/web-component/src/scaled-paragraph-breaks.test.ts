// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createDocument } from 'docx-core';
import { modelToDoc } from './model-adapter';
import { runStylesPlugin } from './run-styles';

describe('scaled paragraph break decorations', () => {
	beforeEach(() => {
		vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
	});
	afterEach(() => vi.restoreAllMocks());
	it('uses whole-word boundaries across marks without changing document text or formatting', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'mixed',
				runs: [
					{ text: 'long', textScalePercent: 50 },
					{ text: 'word next-word あい', textScalePercent: 50, bold: true },
				],
			},
		];
		const doc = modelToDoc(model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, plugins: [runStylesPlugin(() => model)] }),
		});
		try {
			expect(view.dom.textContent).toBe('longword next-word あい');
			expect(view.state.doc.eq(doc)).toBe(true);
			const paragraph = view.dom.querySelector('p')!;
			expect(paragraph.style.whiteSpace).toBe('pre');
			const offsets = [...paragraph.querySelectorAll('wbr')].map((node) => {
				const range = document.createRange();
				range.selectNodeContents(paragraph);
				range.setEndBefore(node);
				return range.toString().length;
			});
			expect(offsets).toEqual([9, 14, 19, 20]);
			expect(paragraph.querySelector('strong')).not.toBeNull();
		} finally {
			view.destroy();
		}
	});
	it('leaves ordinary paragraph wrapping untouched', () => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'plain', runs: [{ text: 'plain text' }] }];
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: modelToDoc(model),
				plugins: [runStylesPlugin(() => model)],
			}),
		});
		try {
			expect(view.dom.querySelector('wbr')).toBeNull();
			expect(view.dom.querySelector('p')!.style.whiteSpace).toBe('');
		} finally {
			view.destroy();
		}
	});
});
