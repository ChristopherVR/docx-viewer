// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, type DocumentModel, type Table } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { at } from './test-support';

function mergedTableModel(): DocumentModel {
	const model = createDocument();
	const table: Table = {
		type: 'table',
		id: 'tbl',
		structureEditable: false,
		grid: [1000, 2000],
		rows: [
			[
				{
					paragraphs: [{ type: 'paragraph', id: 'a', runs: [{ text: 'A' }] }],
					verticalMerge: 'restart',
					shadingFill: '#ABCDEF',
				},
				{
					paragraphs: [{ type: 'paragraph', id: 'b', runs: [{ text: 'B' }] }],
					nestedTables: [{ rows: [[{ text: 'Nested' }]] }],
				},
			],
			[
				{
					paragraphs: [{ type: 'paragraph', id: 'a-continue', runs: [{ text: '' }] }],
					verticalMerge: 'continue',
				},
				{ paragraphs: [{ type: 'paragraph', id: 'd', runs: [{ text: 'D' }] }] },
			],
		],
	};
	model.blocks = [table];
	return model;
}

describe('merged table rendering and structure-preserving edits', () => {
	it('renders a vertical merge as rowspan and omits the continuation cell from the visible doc', () => {
		const model = mergedTableModel();
		const doc = modelToDoc(model);
		const table = doc.firstChild!;
		expect(table.childCount).toBe(2); // two rows
		expect(table.child(0).childCount).toBe(2); // row 0: both cells visible
		expect(table.child(0).child(0).attrs.rowspan).toBe(2);
		expect(table.child(0).child(0).attrs.shadingFill).toBe('#ABCDEF');
		expect(table.child(1).childCount).toBe(1); // row 1: only the non-merged cell
		expect(table.child(1).child(0).firstChild?.textContent).toBe('D');
	});

	it('renders a nested table as a read-only preview node', () => {
		const model = mergedTableModel();
		const doc = modelToDoc(model);
		const cellB = doc.firstChild!.child(0).child(1);
		const preview = cellB.lastChild!;
		expect(preview.type.name).toBe('nestedTablePreview');
		expect(JSON.parse(String(preview.attrs.rowsJson))).toEqual([[{ text: 'Nested' }]]);
	});

	it('edits paragraph text inside a merged table while preserving merges, shading and the nested preview', () => {
		const model = mergedTableModel();
		const doc = modelToDoc(model);
		const state = EditorState.create({ doc });
		// Find the "D" paragraph's text position and insert at its end.
		let pos = -1;
		state.doc.descendants((node, at) => {
			if (node.isText && node.text === 'D') pos = at + 1;
		});
		const tr = state.tr.insertText('-edited', pos);
		const next = docToModel(tr.doc, model);
		const table = next.blocks[0] as Table;
		expect(table.rows).toHaveLength(2);
		const cellAt = (row: number, column: number) => at(at(table.rows, row), column);
		expect(cellAt(0, 0)).toMatchObject({ verticalMerge: 'restart', shadingFill: '#ABCDEF' });
		expect(cellAt(1, 0)).toMatchObject({ verticalMerge: 'continue' });
		expect(cellAt(0, 1).nestedTables).toEqual([{ rows: [[{ text: 'Nested' }]] }]);
		expect(at(at(cellAt(1, 1).paragraphs, 0).runs, 0).text).toBe('D-edited');
	});

	it('round-trips a real merged-table DOCX package end to end through save', async () => {
		const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const xml = `<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl><w:sectPr/></w:body></w:document>`;
		const zip = new JSZip();
		zip.file('word/document.xml', xml);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const doc = modelToDoc(loaded.model);
		const state = EditorState.create({ doc });
		let pos = -1;
		state.doc.descendants((node, at) => {
			if (node.isText && node.text === 'D') pos = at + 1;
		});
		const next = docToModel(state.tr.insertText('!', pos).doc, loaded.model);
		const saved = await JSZip.loadAsync(await loaded.save(next));
		const savedXml = await saved.file('word/document.xml')?.async('string');
		expect(savedXml).toContain('D!');
		expect(savedXml).toContain('<w:vMerge w:val="restart"');
		expect(savedXml).toMatch(/<w:vMerge\s*\/>/);
	});
});
