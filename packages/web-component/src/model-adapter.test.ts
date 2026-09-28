// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx } from '@christophervr/docx-core';
import { DOMParser as ProseMirrorDOMParser } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { insertHardBreak } from './hard-break-command';
import { schema } from './schema';

describe('paragraph line breaks and spacing adapter', () => {
	it('retains an explicit left override instead of treating it as inherited alignment', () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'styled',
			style: 'Centered',
			runs: [{ text: 'text' }],
		};
		const state = EditorState.create({ doc: modelToDoc(model) });
		const next = docToModel(state.tr.setNodeAttribute(0, 'align', 'left').doc, model);
		expect(next.blocks[0]).toMatchObject({ style: 'Centered', align: 'left' });
		expect(next.blocks[0]).not.toBe(model.blocks[0]);
	});
	it('maps model newlines to hard breaks and roundtrips them through DOCX', async () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'line-breaks',
				runs: [{ text: 'First\nSecond\nThird', bold: true, fontFamily: 'Georgia' }],
			},
		];

		const editorDoc = modelToDoc(model);
		const paragraph = editorDoc.firstChild!;
		expect(paragraph.content.content.map((node) => node.type.name)).toEqual([
			'text',
			'hardBreak',
			'text',
			'hardBreak',
			'text',
		]);
		expect(paragraph.child(1).marks.map((mark) => mark.type.name)).toEqual(['bold', 'font']);
		expect(editorDoc.textBetween(0, editorDoc.content.size, '\n')).toBe('First\nSecond\nThird');

		const roundtripped = docToModel(editorDoc, model);
		expect(roundtripped.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: 'First\nSecond\nThird', bold: true, fontFamily: 'Georgia' }],
		});
		const reopened = await loadDocx(await saveDocx(roundtripped));
		expect(reopened.model.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: 'First\nSecond\nThird', bold: true, fontFamily: 'Georgia' }],
		});
	});

	it('parses pasted HTML breaks as hard breaks within one paragraph', () => {
		const pasted = ProseMirrorDOMParser.fromSchema(schema).parse(
			new DOMParser().parseFromString('<p>one<br>two</p>', 'text/html').body,
		);
		expect(pasted.firstChild?.childCount).toBe(3);
		expect(pasted.firstChild?.child(1).type.name).toBe('hardBreak');
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'pasted', runs: [{ text: 'one\ntwo' }] }];
		expect(docToModel(pasted, model).blocks[0]).toMatchObject({
			runs: [{ text: 'one\ntwo' }],
		});
	});

	it('inserts a hard break command without splitting the paragraph', () => {
		const paragraph = schema.nodes.paragraph.create({ id: 'command' }, [schema.text('ab')]);
		const doc = schema.nodes.doc.create(null, [paragraph]);
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 2) });
		let changed = state;
		expect(
			insertHardBreak(state, (transaction) => {
				changed = state.apply(transaction);
			}),
		).toBe(true);
		expect(changed.doc.firstChild?.childCount).toBe(3);
		expect(changed.doc.firstChild?.child(1).type.name).toBe('hardBreak');
		expect(changed.doc.childCount).toBe(1);
	});

	it('uses Word line spacing as a line multiplier by default', () => {
		const paragraph = schema.nodes.paragraph.create({ lineSpacingTwips: 360 });
		const dom = schema.nodes.paragraph.spec.toDOM?.(paragraph);
		expect(dom).toEqual([
			'p',
			{ style: 'text-align:left;line-height:1.8', dir: null, 'data-id': '' },
			0,
		]);
	});

	it('preserves a line spacing rule even if its value is absent during another edit', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'rule-only',
				runs: [{ text: 'Keep rule' }],
				lineSpacingRule: 'atLeast',
			},
		];
		const editorDoc = modelToDoc(model);
		const changed = schema.nodes.doc.create({ ...editorDoc.attrs }, [
			schema.nodes.paragraph.create(
				{ ...editorDoc.firstChild!.attrs, align: 'center' },
				editorDoc.firstChild!.content,
			),
		]);
		expect(docToModel(changed, model).blocks[0]).toMatchObject({
			align: 'center',
			lineSpacingRule: 'atLeast',
		});
	});
});
