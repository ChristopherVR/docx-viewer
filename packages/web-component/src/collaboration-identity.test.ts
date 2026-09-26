import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import {
	createCollaborationIdGenerator,
	repairCollaborativeDocumentIds,
} from './collaboration-identity';
import { schema } from './schema';

describe('collaboration document identities', () => {
	it('namespaces inserted IDs by client and repairs duplicate split paragraph IDs', () => {
		const doc = schema.nodes.doc.create(null, [
			schema.nodes.paragraph.create({ id: 'original' }, [schema.text('split')]),
		]);
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 3) });
		const split = state.tr.split(3);
		const splitState = state.apply(split);
		const left = repairCollaborativeDocumentIds(splitState, 'peer/left')!;
		const right = repairCollaborativeDocumentIds(splitState, 'peer:right')!;
		const leftDoc = splitState.apply(left).doc;
		const rightDoc = splitState.apply(right).doc;
		const ids = (value: typeof doc) => {
			const result: string[] = [];
			value.descendants((node) => {
				if (node.type === schema.nodes.paragraph) result.push(node.attrs.id);
			});
			return result;
		};
		expect(new Set(ids(leftDoc)).size).toBe(2);
		expect(new Set(ids(rightDoc)).size).toBe(2);
		expect(ids(leftDoc)[1]).not.toBe(ids(rightDoc)[1]);
		expect(ids(leftDoc)[1]).toMatch(/^dve-paragraph-/);
	});

	it('provides a reusable per-client ID generator for table and cell insertions', () => {
		const left = createCollaborationIdGenerator('left');
		const right = createCollaborationIdGenerator('right');
		const leftIds = [left('table'), left('cell'), left('cell')];
		const rightIds = [right('table'), right('cell'), right('cell')];
		expect(new Set(leftIds).size).toBe(leftIds.length);
		expect(new Set(rightIds).size).toBe(rightIds.length);
		expect(leftIds.every((id) => !rightIds.includes(id))).toBe(true);
	});

	it('repairs table and cell paragraph IDs from the persistent client generator', () => {
		const table = schema.nodes.table.create(null, [
			schema.nodes.tableRow.create(null, [
				schema.nodes.tableCell.create(null, schema.nodes.paragraph.create(null)),
			]),
		]);
		const doc = schema.nodes.doc.create(null, [table]);
		const clientIds = createCollaborationIdGenerator('table-peer');
		const state = EditorState.create({ doc });
		const transaction = repairCollaborativeDocumentIds(state, 'table-peer', clientIds)!;
		const repaired = state.apply(transaction).doc;
		const tableId = repaired.firstChild!.attrs.id;
		let paragraphId: string | null = null;
		repaired.descendants((node) => {
			if (node.type === schema.nodes.paragraph) paragraphId = node.attrs.id;
		});
		expect(paragraphId).not.toBeNull();
		expect(tableId).toMatch(/^dve-table-/);
		expect(paragraphId).toMatch(/^dve-paragraph-/);
		expect(tableId).not.toBe(paragraphId);
		const next = clientIds('table');
		expect(next).not.toBe(tableId);
	});
});
