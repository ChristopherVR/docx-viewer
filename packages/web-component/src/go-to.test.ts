// @vitest-environment jsdom
import { createDocument } from 'docx-core';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { goToTarget, goToTargets, selectNextObject } from './go-to';
import { openGoToPanel } from './go-to-panel';
import { createRibbon } from './ribbon';
import { schema } from './schema';

beforeAll(() => {
	// jsdom has no layout; ProseMirror asks the DOM range for rectangles when it scrolls.
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});

afterEach(() => (document.body.innerHTML = ''));

const model = createDocument();

function editor() {
	const image = (altText: string | null = null) =>
		schema.nodes.image!.create({
			partName: 'word/media/a.png',
			altText,
			widthPx: 10,
			heightPx: 10,
		});
	const cell = schema.nodes.tableCell!.create(
		{},
		schema.nodes.paragraph!.create({ id: 'c' }, schema.text('Cell text')),
	);
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'h1', style: 'Heading1' }, schema.text('Introduction')),
		schema.nodes.paragraph!.create({ id: 'body', bookmarks: ['Intro', '_Toc123'] }, [
			schema.text('Body '),
			image('A chart'),
			schema.text(' more'),
		]),
		schema.nodes.paragraph!.create({ id: 'h2', style: 'Heading2' }, schema.text('Details')),
		schema.nodes.table!.create({ id: 't' }, schema.nodes.tableRow!.create({}, cell)),
		schema.nodes.paragraph!.create({ id: 'end' }, [image()]),
	]);
	const host = document.createElement('div');
	document.body.append(host);
	return new EditorView(host, { state: EditorState.create({ doc, schema }) });
}

describe('goToTargets', () => {
	it('lists headings by style, indented by level', () => {
		const view = editor();
		const labels = goToTargets(view.state.doc, model, 'heading').map((t) => t.label);
		expect(labels).toEqual(['Introduction', '  Details']);
	});

	it("lists user bookmarks but not Word's hidden ones", () => {
		expect(goToTargets(editor().state.doc, model, 'bookmark').map((t) => t.label)).toEqual([
			'Intro',
		]);
	});

	it('lists tables and pictures', () => {
		const view = editor();
		expect(goToTargets(view.state.doc, model, 'table').map((t) => t.label)).toEqual([
			'1: Cell text',
		]);
		expect(goToTargets(view.state.doc, model, 'graphic').map((t) => t.label)).toEqual([
			'1: A chart',
			'2',
		]);
	});
});

describe('goToTarget', () => {
	it('places the caret in a heading, a table cell and on a picture', () => {
		const view = editor();
		const [intro] = goToTargets(view.state.doc, model, 'heading');
		expect(goToTarget(view, intro!)).toBe(true);
		expect(view.state.selection.$from.parent.textContent).toBe('Introduction');
		goToTarget(view, goToTargets(view.state.doc, model, 'table')[0]!);
		expect(view.state.selection.$from.parent.textContent).toBe('Cell text');
		goToTarget(view, goToTargets(view.state.doc, model, 'graphic')[1]!);
		expect(view.state.selection).toBeInstanceOf(NodeSelection);
	});

	it('refuses a stale position', () => {
		const view = editor();
		expect(goToTarget(view, { label: 'gone', pos: 9999 })).toBe(false);
		expect(goToTarget(view, { label: 'not a picture', pos: 1, node: true })).toBe(false);
	});
});

describe('selectNextObject', () => {
	it('walks the pictures and wraps around', () => {
		const view = editor();
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		expect(selectNextObject(view)).toBe(true);
		const first = view.state.selection.from;
		selectNextObject(view);
		const second = view.state.selection.from;
		expect(second).toBeGreaterThan(first);
		selectNextObject(view);
		expect(view.state.selection.from).toBe(first);
	});

	it('does nothing without pictures', () => {
		const doc = schema.node('doc', null, [
			schema.nodes.paragraph!.create({ id: 'p' }, schema.text('x')),
		]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
		expect(selectNextObject(view)).toBe(false);
	});
});

describe('Go To panel and Editing controls', () => {
	it('lists targets for the chosen kind and jumps to one', () => {
		const view = editor();
		const anchor = document.createElement('button');
		document.body.append(anchor);
		let closed = 0;
		openGoToPanel(anchor, view, model, () => closed++);
		const panel = document.querySelector('.go-to-panel')!;
		const items = () => [...panel.querySelectorAll<HTMLButtonElement>('[role="option"]')];
		expect(items().map((b) => b.textContent)).toEqual(['Introduction', '  Details']);
		const kind = panel.querySelector<HTMLSelectElement>('select')!;
		kind.value = 'table';
		kind.dispatchEvent(new Event('change'));
		expect(items()).toHaveLength(1);
		items()[0]!.click();
		expect(view.state.selection.$from.parent.textContent).toBe('Cell text');
		expect(closed).toBe(1);
		kind.value = 'bookmark';
		kind.dispatchEvent(new Event('change'));
		kind.value = 'heading';
		kind.dispatchEvent(new Event('change'));
		expect(items()).toHaveLength(2);
	});

	it('says when there is nothing to go to', () => {
		const doc = schema.node('doc', null, [schema.nodes.paragraph!.create({ id: 'p' })]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
		const anchor = document.createElement('button');
		document.body.append(anchor);
		openGoToPanel(anchor, view, model, () => {});
		expect(document.querySelector('.go-to-list')!.textContent).toBe('Nothing to go to.');
	});

	it('offers Find, Replace and Select with caret menus that emit their actions', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const find = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Find options"]')!;
		find.value = 'goto';
		find.dispatchEvent(new Event('change'));
		const select = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Select options"]')!;
		select.value = 'objects';
		select.dispatchEvent(new Event('change'));
		expect(seen).toEqual([{ type: 'goTo' }, { type: 'selectObjects' }]);
		expect(find.selectedIndex).toBe(-1);
		expect(ribbon.querySelectorAll('.ribbon-split-inline')).toHaveLength(2);
	});
});
