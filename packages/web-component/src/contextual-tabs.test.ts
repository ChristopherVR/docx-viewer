// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { describe, expect, it } from 'vitest';
import { selectionInTable, syncContextualTabs } from './contextual-tabs';
import { createRibbon } from './ribbon';
import { schema } from './schema';

function tableState() {
	const cell = schema.nodes.tableCell!.create(
		{},
		schema.nodes.paragraph!.create({ id: 'c' }, schema.text('cell')),
	);
	const table = schema.nodes.table!.create({ id: 't' }, schema.nodes.tableRow!.create({}, cell));
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'a' }, schema.text('before')),
		table,
	]);
	return EditorState.create({ doc, schema });
}

describe('contextual Table tab', () => {
	it('is hidden in a new ribbon and tinted as a contextual tab', () => {
		const tab = createRibbon().querySelector<HTMLElement>('#dve-tab-table')!;
		expect(tab.hidden).toBe(true);
		expect(tab.hasAttribute('data-contextual')).toBe(true);
	});

	it('appears while the selection is in a table and leaves when it moves out', () => {
		const ribbon = createRibbon();
		let state = tableState();
		const inCell = state.doc.resolve(state.doc.content.size - 4).pos;
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, inCell)));
		expect(selectionInTable(state)).toBe(true);
		syncContextualTabs(ribbon, state);
		expect(ribbon.querySelector<HTMLElement>('#dve-tab-table')!.hidden).toBe(false);
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
		syncContextualTabs(ribbon, state);
		expect(ribbon.querySelector<HTMLElement>('#dve-tab-table')!.hidden).toBe(true);
	});

	it('returns to Home when the table tab was open as the selection leaves', () => {
		const ribbon = createRibbon();
		let state = tableState();
		const inCell = state.doc.content.size - 4;
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, inCell)));
		syncContextualTabs(ribbon, state);
		ribbon.querySelector<HTMLButtonElement>('#dve-tab-table')!.click();
		expect(ribbon.querySelector('#dve-panel-table')!.hasAttribute('hidden')).toBe(false);
		state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
		syncContextualTabs(ribbon, state);
		expect(ribbon.querySelector('#dve-tab-home')!.getAttribute('aria-selected')).toBe('true');
		expect(ribbon.querySelector('#dve-panel-table')!.hasAttribute('hidden')).toBe(true);
	});

	it('skips the hidden tab when arrowing through the tab strip', () => {
		const ribbon = createRibbon();
		const view = ribbon.querySelector<HTMLButtonElement>('#dve-tab-view')!;
		view.focus();
		view.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(ribbon.querySelector('#dve-tab-home')!.getAttribute('aria-selected')).toBe('true');
	});
});
