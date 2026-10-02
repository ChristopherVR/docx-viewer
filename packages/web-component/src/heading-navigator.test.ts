// @vitest-environment jsdom
import { createDocument, type DocumentModel } from 'docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it, vi } from 'vitest';
import { HeadingNavigator } from './heading-navigator';
import { modelToDoc } from './model-adapter';
import { schema } from './schema';

function setup() {
	let model: DocumentModel = {
		...createDocument(),
		blocks: [
			{ type: 'paragraph', id: 'a', style: 'Heading1', runs: [{ text: 'Chapter' }] },
			{ type: 'paragraph', id: 'b', style: 'Child', runs: [{ text: 'Topic' }] },
			{ type: 'paragraph', id: 'body', runs: [{ text: 'Body text' }] },
			{ type: 'paragraph', id: 'c', style: 'Heading1', runs: [{ text: 'Next chapter' }] },
		],
		paragraphStyles: {
			docDefaults: {},
			warnings: [],
			styles: {
				Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} },
				Heading2: { id: 'Heading2', name: 'heading 2', formatting: {} },
				Child: { id: 'Child', name: 'Custom', basedOn: 'Heading2', formatting: {} },
			},
		},
	};
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ schema, doc: modelToDoc(model) }),
	});
	const goTo = vi.fn();
	const close = vi.fn();
	const nav = new HeadingNavigator({ model: () => model, view: () => view, goTo, close });
	document.body.append(nav.element);
	nav.setOpen(true);
	return {
		nav,
		view,
		goTo,
		close,
		setModel: (value: DocumentModel) => {
			model = value;
		},
		model,
		items: () => [...nav.element.querySelectorAll<HTMLElement>('[role=treeitem]')],
	};
}

describe('heading navigation', () => {
	it('resolves inherited headings, excludes body text, collapses descendants and navigates with keys', () => {
		const { nav, view, goTo, items } = setup();
		expect(items().map((item) => item.textContent?.replace('▾', ''))).toEqual([
			'Chapter',
			'Topic',
			'Next chapter',
		]);
		expect(items()[1]!.getAttribute('aria-level')).toBe('2');
		items()[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true }));
		expect(items()[1]!.hidden).toBe(true);
		items()[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		expect(document.activeElement).toBe(items()[2]);
		items()[2]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(goTo).toHaveBeenCalledWith('c');
		items()[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(items()[1]!.hidden).toBe(false);
		nav.element.remove();
		view.destroy();
	});
	it('tracks the selection, refreshed text and locale, and explains an empty outline', () => {
		const { nav, view, setModel, model, items } = setup();
		view.dispatch(
			view.state.tr.setSelection(
				TextSelection.create(view.state.doc, view.state.doc.child(0).nodeSize + 2),
			),
		);
		nav.sync('en');
		expect(items()[1]!.getAttribute('aria-selected')).toBe('true');
		setModel({
			...model,
			blocks: [{ type: 'paragraph', id: 'new', style: 'Heading1', runs: [{ text: 'Changed' }] }],
		});
		nav.sync('fr');
		expect(nav.element.getAttribute('aria-label')).toBe('Volet de navigation');
		expect(items()[0]!.textContent).toBe('Changed');
		setModel({ ...model, blocks: [] });
		nav.sync('en');
		expect(items()).toHaveLength(0);
		expect(nav.element.querySelector('[role=status]')!.textContent).toContain(
			'Apply a heading style',
		);
		nav.element.remove();
		view.destroy();
	});
});
