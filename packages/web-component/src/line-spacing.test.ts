// @vitest-environment jsdom
import { signedTwips, twips } from 'docx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument } from 'docx-core';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { lineSpacingLabel, lineSpacingValue } from './line-spacing';
import { at, tableAt } from './test-support';

afterEach(() => document.body.replaceChildren());

function mountTableEditor() {
	registerDocxEditor();
	const model = createDocument();
	model.blocks = [
		{
			type: 'table',
			id: 'spacing-table',
			rows: [
				[
					{
						paragraphs: [
							{
								type: 'paragraph',
								id: 'first-cell',
								runs: [{ text: 'first cell' }],
								lineSpacingTwips: signedTwips(301),
								lineSpacingRule: 'exact',
								spacingBeforeTwips: twips(80),
								align: 'center',
							},
						],
					},
					{
						paragraphs: [
							{
								type: 'paragraph',
								id: 'second-cell',
								runs: [{ text: 'second cell' }],
								lineSpacingTwips: signedTwips(360),
								lineSpacingRule: 'atLeast',
								spacingAfterTwips: twips(100),
								indentLeftTwips: signedTwips(120),
							},
						],
					},
				],
			],
		},
	];
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	const view = (editor as unknown as { view: EditorView }).view;
	const textRanges: Array<{ from: number; to: number }> = [];
	view.state.doc.descendants((node, pos) => {
		if (node.isText) textRanges.push({ from: pos, to: pos + node.nodeSize });
	});
	return { editor, view, textRanges };
}

function chooseSpacing(editor: DocxEditorElement, value: string) {
	const select = editor.shadowRoot!.querySelector<HTMLSelectElement>(
		'[aria-label="Line spacing"]',
	)!;
	select.value = value;
	select.dispatchEvent(new Event('change', { bubbles: true }));
}

function documentModel(editor: DocxEditorElement) {
	const model = editor.documentModel;
	if (!model) throw new Error('Expected an editor document');
	return model;
}

describe('line spacing ribbon control', () => {
	it('labels malformed imported rules without presenting them as inherited', () => {
		const value = lineSpacingValue({ lineSpacingRule: 'exact' });
		expect(value).toBe('rule-only:exact');
		expect(lineSpacingLabel(value)).toBe('Exact rule (no amount)');
	});

	it('shows imported exact and at-least values, then reports a mixed table-cell selection', () => {
		const { editor, view, textRanges } = mountTableEditor();
		const control = editor.shadowRoot!.querySelector<HTMLSelectElement>(
			'[aria-label="Line spacing"]',
		)!;

		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, at(textRanges, 0).from, at(textRanges, 0).to),
			),
		);
		expect(control.value).toBe('exact:301');
		expect([...control.options].find((option) => option.value === 'exact:301')?.textContent).toBe(
			'Exact 15.05 pt',
		);

		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, at(textRanges, 0).from, at(textRanges, 1).to),
			),
		);
		expect(control.value).toBe('mixed');
		expect([...control.options].find((option) => option.value === 'atLeast:360')?.textContent).toBe(
			'At least 18 pt',
		);
		editor.remove();
	});

	it('updates every selected table-cell paragraph, preserves other attrs, supports undo and redo, and clears to inherit', () => {
		const { editor, view, textRanges } = mountTableEditor();
		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, at(textRanges, 0).from, at(textRanges, 1).to),
			),
		);
		chooseSpacing(editor, 'auto:360');
		const table = tableAt(documentModel(editor).blocks, 0);
		const [firstRow = []] = table.rows;
		expect(at(at(firstRow, 0).paragraphs, 0)).toMatchObject({
			lineSpacingTwips: 360,
			lineSpacingRule: 'auto',
			spacingBeforeTwips: 80,
			align: 'center',
		});
		expect(at(at(firstRow, 1).paragraphs, 0)).toMatchObject({
			lineSpacingTwips: 360,
			lineSpacingRule: 'auto',
			spacingAfterTwips: 100,
			indentLeftTwips: 120,
		});

		chooseSpacing(editor, 'auto:480');
		editor.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Undo"]')!.click();
		expect(documentModel(editor).blocks[0]).toMatchObject({
			rows: [
				[
					{ paragraphs: [{ lineSpacingTwips: 360, lineSpacingRule: 'auto' }] },
					{ paragraphs: [{ lineSpacingTwips: 360, lineSpacingRule: 'auto' }] },
				],
			],
		});
		editor.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Redo"]')!.click();
		expect(documentModel(editor).blocks[0]).toMatchObject({
			rows: [
				[
					{ paragraphs: [{ lineSpacingTwips: 480, lineSpacingRule: 'auto' }] },
					{ paragraphs: [{ lineSpacingTwips: 480, lineSpacingRule: 'auto' }] },
				],
			],
		});

		chooseSpacing(editor, 'inherit');
		const cleared = tableAt(documentModel(editor).blocks, 0);
		const clearedRow = at(cleared.rows, 0);
		expect(clearedRow.map((cell) => at(cell.paragraphs, 0))).toMatchObject([
			{ spacingBeforeTwips: 80, align: 'center' },
			{ spacingAfterTwips: 100, indentLeftTwips: 120 },
		]);
		expect(
			clearedRow
				.flatMap((cell) => cell.paragraphs)
				.every((paragraph) => paragraph.lineSpacingTwips == null),
		).toBe(true);
		editor.readOnly = true;
		const before = structuredClone(documentModel(editor));
		const control = editor.shadowRoot!.querySelector<HTMLSelectElement>(
			'[aria-label="Line spacing"]',
		)!;
		expect(control.disabled).toBe(true);
		chooseSpacing(editor, 'auto:720');
		expect(documentModel(editor)).toEqual(before);
		editor.remove();
	});
});
