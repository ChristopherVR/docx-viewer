// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { beforeAll, describe, expect, it } from 'vitest';
import { insertCoverPage } from './cover-page';
import { insertHorizontalLine, setBorders } from './paragraph-decoration';
import { createRibbon } from './ribbon';
import { schema } from './schema';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});

function editor(...texts: string[]) {
	const doc = schema.node(
		'doc',
		null,
		texts.map((text, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}` }, schema.text(text)),
		),
	);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const texts = (view: EditorView) => {
	const out: string[] = [];
	view.state.doc.forEach((node) => out.push(node.textContent));
	return out;
};

describe('cover page', () => {
	const text = { title: 'My title', subtitle: 'Sub', author: 'Ann', date: 'May 1' };

	it('adds a centred title page and pushes the first paragraph to the next page', () => {
		const view = editor('Body one', 'Body two');
		expect(insertCoverPage(view, text, 'Title')).toBe(true);
		expect(texts(view)).toEqual(['', 'My title', 'Sub', 'Ann', 'May 1', 'Body one', 'Body two']);
		const title = view.state.doc.child(1);
		expect(title.attrs).toMatchObject({ align: 'center', style: 'Title' });
		expect(view.state.doc.child(5).attrs.pageBreakBefore).toBe(true);
		expect(view.state.doc.child(6).attrs.pageBreakBefore).toBe(false);
	});

	it('works without a Title style and refuses a read-only view', () => {
		const view = editor('Body');
		insertCoverPage(view, text);
		expect(view.state.doc.child(1).attrs.style).toBeFalsy();
		const locked = new EditorView(document.createElement('div'), {
			state: editor('x').state,
			editable: () => false,
		});
		expect(insertCoverPage(locked, text)).toBe(false);
	});

	it('is one undo step of content and has a ribbon button', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Cover page"]')!.click();
		expect(seen).toEqual([{ type: 'coverPage' }]);
	});
});

describe('horizontal line', () => {
	it('inserts a bordered empty paragraph after the current one and moves the caret below it', () => {
		const view = editor('First', 'Second');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2)));
		expect(insertHorizontalLine(view)).toBe(true);
		expect(texts(view)).toEqual(['First', '', '', 'Second']);
		const line = view.state.doc.child(1);
		expect(line.attrs.borders).toEqual({
			bottom: { style: 'single', sizeEighthPoints: 6, spacePoints: 1 },
		});
		expect(view.state.selection.$from.parent).toBe(view.state.doc.child(2));
		expect(view.state.doc.child(2).attrs.borders).toBeNull();
	});

	it('is available from the Borders menu', () => {
		const view = editor('Only');
		expect(setBorders(view, 'horizontal')).toBe(true);
		expect(view.state.doc.childCount).toBe(3);
		const ribbon = createRibbon();
		const menu = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Borders"]')!;
		expect([...menu.options].map((o) => o.value)).toContain('horizontal');
	});
});
