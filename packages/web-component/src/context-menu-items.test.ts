// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import {
	clampToViewport,
	computeMenuItems,
	menuStateFromView,
	type MenuState,
} from './context-menu-items';

const base: MenuState = {
	readOnly: false,
	hasSelection: true,
	onLink: false,
	inTable: false,
	canReadClipboard: true,
	canTable: () => true,
};
const ids = (state: Partial<MenuState>) =>
	computeMenuItems({ ...base, ...state }).map((item) => item.id);
const disabled = (state: Partial<MenuState>) =>
	computeMenuItems({ ...base, ...state })
		.filter((item) => item.disabled)
		.map((item) => item.id);

describe('context menu items', () => {
	it('offers clipboard, link and comment items outside tables', () => {
		expect(ids({})).toEqual(['cut', 'copy', 'paste', 'link-insert', 'comment-add']);
	});

	it('swaps the insert link item for edit/remove on a link', () => {
		expect(ids({ onLink: true })).toEqual([
			'cut',
			'copy',
			'paste',
			'link-edit',
			'link-remove',
			'comment-add',
		]);
	});

	it('adds table actions only inside a table', () => {
		expect(ids({}).some((id) => id.startsWith('table-'))).toBe(false);
		expect(ids({ inTable: true }).filter((id) => id.startsWith('table-'))).toEqual([
			'table-rowBefore',
			'table-rowAfter',
			'table-deleteRow',
			'table-columnBefore',
			'table-columnAfter',
			'table-deleteColumn',
			'table-deleteTable',
		]);
	});

	it('disables cut and copy without a selection', () => {
		expect(disabled({ hasSelection: false })).toEqual(['cut', 'copy']);
	});

	it('disables everything but copy when read-only', () => {
		const readOnly = computeMenuItems({ ...base, readOnly: true, inTable: true });
		expect(readOnly.filter((item) => !item.disabled).map((item) => item.id)).toEqual(['copy']);
	});

	it('disables paste honestly, with a reason, when the clipboard cannot be read', () => {
		const paste = computeMenuItems({ ...base, canReadClipboard: false }).find(
			(item) => item.id === 'paste',
		)!;
		expect(paste.disabled).toBe(true);
		expect(paste.reason).toBe('menu.pasteUnavailable');
	});

	it('disables table items the table does not allow', () => {
		const items = computeMenuItems({
			...base,
			inTable: true,
			canTable: (command) => command === 'deleteTable',
		});
		expect(items.filter((item) => item.id.startsWith('table-') && !item.disabled)).toHaveLength(1);
	});

	it('separates groups', () => {
		const items = computeMenuItems({ ...base, inTable: true });
		expect(items.filter((item) => item.separatorBefore).map((item) => item.id)).toEqual([
			'link-insert',
			'comment-add',
			'table-rowBefore',
		]);
	});
});

describe('context menu state from a view', () => {
	const table = schema.nodes.table.createAndFill()!;
	const doc = schema.node('doc', undefined, [
		schema.node('paragraph', undefined, [schema.text('outside')]),
		table,
	]);
	const view = (position: number, editable = true) => {
		const state = EditorState.create({ doc });
		return new EditorView(document.createElement('div'), {
			state: state.apply(state.tr.setSelection(TextSelection.create(state.doc, position))),
			editable: () => editable,
		});
	};

	it('detects the table from the selection and reflects read-only', () => {
		expect(menuStateFromView(view(2), false).inTable).toBe(false);
		const inside = doc.child(0).nodeSize + 3;
		expect(menuStateFromView(view(inside), false).inTable).toBe(true);
		expect(menuStateFromView(view(inside, false), false).readOnly).toBe(true);
		expect(menuStateFromView(view(inside), true).readOnly).toBe(true);
	});
});

describe('viewport clamping', () => {
	const size = { width: 200, height: 300 };
	const viewport = { width: 1000, height: 600 };
	it('keeps the requested position when it fits', () => {
		expect(clampToViewport(100, 100, size, viewport)).toEqual({ left: 100, top: 100 });
	});
	it('shifts the menu back inside on the right and bottom edges', () => {
		expect(clampToViewport(950, 590, size, viewport)).toEqual({ left: 796, top: 296 });
	});
	it('never goes above or left of the margin, even when larger than the viewport', () => {
		expect(clampToViewport(-20, -5, size, viewport)).toEqual({ left: 4, top: 4 });
		expect(clampToViewport(10, 10, { width: 2000, height: 900 }, viewport)).toEqual({
			left: 4,
			top: 4,
		});
	});
});
