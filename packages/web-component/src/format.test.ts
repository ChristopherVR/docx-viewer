import { describe, expect, it } from 'vitest';
import { createDocument, isWordHighlightToken } from '@christophervr/docx-core';
import { modelToDoc, docToModel } from './model-adapter';
import { schema, wordHighlightColors } from './schema';

describe('Word inline formatting', () => {
	it('preserves strike, named highlight, and vertical alignment through the editor adapter', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'format-test',
				runs: [
					{ text: 'marked', strike: true, highlight: 'yellow', verticalAlign: 'superscript' },
					{ text: 'sub', verticalAlign: 'subscript' },
				],
			},
		];

		const editorDoc = modelToDoc(model);
		const updated = docToModel(editorDoc, model);
		const runs = updated.blocks[0].type === 'paragraph' ? updated.blocks[0].runs : [];
		expect(runs).toEqual([
			{ text: 'marked', strike: true, highlight: 'yellow', verticalAlign: 'superscript' },
			{ text: 'sub', verticalAlign: 'subscript' },
		]);
	});

	it('renders all supported Word highlight tokens with their palette colors', () => {
		for (const [color, cssColor] of Object.entries(wordHighlightColors)) {
			const mark = schema.marks.highlight.create({ color });
			expect(schema.marks.highlight.spec.toDOM?.(mark, false)).toEqual([
				'span',
				{ style: `background-color:${cssColor}` },
				0,
			]);
		}
	});

	it('accepts canonical OOXML highlight names and rejects VBA-only aliases', () => {
		expect(isWordHighlightToken('cyan')).toBe(true);
		expect(isWordHighlightToken('darkGreen')).toBe(true);
		expect(isWordHighlightToken('turquoise')).toBe(false);
		expect(isWordHighlightToken('gray25')).toBe(false);
		expect(wordHighlightColors.cyan).toBe('#00ffff');
	});

	it('preserves a parser-marked structurally unsafe table through the adapter', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'table',
				id: 'complex-table',
				rows: [[{ paragraphs: [{ type: 'paragraph', id: 'cell-p', runs: [{ text: 'cell' }] }] }]],
				structureEditable: false,
			} as (typeof model.blocks)[number],
		];
		const updated = docToModel(modelToDoc(model), model);
		expect(updated.blocks[0]).toMatchObject({ type: 'table', structureEditable: false });
	});

	it('keeps superscript and subscript mutually exclusive', () => {
		const text = schema.text('x', [schema.marks.verticalAlign.create({ value: 'superscript' })]);
		const changed = text.mark([
			...text.marks.filter((mark) => mark.type.name !== 'verticalAlign'),
			schema.marks.verticalAlign.create({ value: 'subscript' }),
		]);
		expect(changed.marks.filter((mark) => mark.type.name === 'verticalAlign')).toHaveLength(1);
		expect(changed.marks.find((mark) => mark.type.name === 'verticalAlign')?.attrs.value).toBe(
			'subscript',
		);
	});
});
