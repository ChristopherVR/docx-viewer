// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { NumberingCatalog } from '@christophervr/docx-core';
import { schema } from './schema';
import {
	changeListLevel,
	exitListOnEmptyEnter,
	indentListItem,
	listKindOf,
	outdentListItem,
	removeList,
	selectionIsListKind,
	toggleList,
} from './list-commands';

function catalog(): NumberingCatalog {
	return {
		abstractNums: {
			'0': {
				id: '0',
				levels: {
					0: { level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', suffix: 'tab' },
					1: { level: 1, start: 1, numFmt: 'lowerLetter', lvlText: '%2)', suffix: 'tab' },
				},
			},
			'1': {
				id: '1',
				levels: { 0: { level: 0, start: 1, numFmt: 'bullet', lvlText: '•', suffix: 'tab' } },
			},
		},
		nums: { '1': { id: '1', abstractNumId: '0' }, '2': { id: '2', abstractNumId: '1' } },
		warnings: [],
	};
}

function viewWithParagraphs(attrsList: Record<string, unknown>[]): EditorView {
	const doc = schema.nodes.doc.create(
		null,
		attrsList.map((attrs, index) =>
			schema.node('paragraph', { id: `p${index}`, ...attrs }, schema.text('x')),
		),
	);
	const dom = document.createElement('div');
	return new EditorView(dom, { state: EditorState.create({ doc }) });
}

describe('list-commands', () => {
	it('classifies list kind from numId/ilvl against the catalog', () => {
		const view = viewWithParagraphs([{ numId: 1, ilvl: 0 }, { numId: 2, ilvl: 0 }, {}]);
		const numbered = view.state.doc.child(0);
		const bulleted = view.state.doc.child(1);
		const plain = view.state.doc.child(2);
		expect(listKindOf(numbered, catalog())).toBe('decimal');
		expect(listKindOf(bulleted, catalog())).toBe('bullet');
		expect(listKindOf(plain, catalog())).toBeNull();
	});

	it('detects a uniform selection kind across paragraphs', () => {
		const view = viewWithParagraphs([
			{ numId: 1, ilvl: 0 },
			{ numId: 1, ilvl: 1 },
		]);
		const tr = view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		);
		view.updateState(view.state.apply(tr));
		expect(selectionIsListKind(view, 'decimal', catalog())).toBe(true);
		expect(selectionIsListKind(view, 'bullet', catalog())).toBe(false);
	});

	it('applies a fresh numId to a plain paragraph and removes it on toggling off', () => {
		const view = viewWithParagraphs([{}]);
		toggleList(view, false, () => 7);
		expect(view.state.doc.child(0).attrs).toMatchObject({ numId: 7, ilvl: 0 });
		toggleList(view, true, () => {
			throw new Error('should not mint a new numId when already this kind');
		});
		expect(view.state.doc.child(0).attrs.numId).toBeNull();
	});

	it('removes numbering only from paragraphs that currently have it', () => {
		const view = viewWithParagraphs([{ numId: 1, ilvl: 0 }, {}]);
		const removed = removeList(view);
		expect(removed).toBe(true);
		expect(view.state.doc.child(0).attrs.numId).toBeNull();
	});

	it('returns false and makes no change when nothing in the selection is a list', () => {
		const view = viewWithParagraphs([{}]);
		expect(removeList(view)).toBe(false);
		expect(changeListLevel(view, 1)).toBe(false);
	});

	it('changes list level clamped between 0 and 8', () => {
		const view = viewWithParagraphs([{ numId: 1, ilvl: 0 }]);
		changeListLevel(view, -1);
		expect(view.state.doc.child(0).attrs.ilvl).toBe(0);
		for (let i = 0; i < 10; i++) changeListLevel(view, 1);
		expect(view.state.doc.child(0).attrs.ilvl).toBe(8);
	});

	it('Tab/Shift+Tab change level only inside a list item', () => {
		const view = viewWithParagraphs([{ numId: 1, ilvl: 0 }, {}]);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		expect(indentListItem(view.state, view.dispatch)).toBe(true);
		expect(view.state.doc.child(0).attrs.ilvl).toBe(1);
		expect(outdentListItem(view.state, view.dispatch)).toBe(true);
		expect(view.state.doc.child(0).attrs.ilvl).toBe(0);
		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, view.state.doc.content.size - 1),
			),
		);
		expect(indentListItem(view.state, view.dispatch)).toBe(false);
	});

	it('Enter on an empty list item ends the list instead of creating a new item', () => {
		const doc = schema.node('doc', null, [
			schema.node('paragraph', { id: 'p0', numId: 1, ilvl: 0 }),
		]);
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
		let applied = false;
		expect(
			exitListOnEmptyEnter(state, (tr) => {
				applied = true;
				const next = state.apply(tr);
				expect(next.doc.child(0).attrs.numId).toBeNull();
			}),
		).toBe(true);
		expect(applied).toBe(true);
	});

	it('does not intercept Enter on a non-empty list item', () => {
		const doc = schema.node('doc', null, [
			schema.node('paragraph', { id: 'p0', numId: 1, ilvl: 0 }, schema.text('text')),
		]);
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
		expect(exitListOnEmptyEnter(state, () => {})).toBe(false);
	});
});
