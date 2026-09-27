// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	createDocument,
	loadDocx,
	saveDocx,
	type DocumentModel,
	type Table,
	type TableStyleCatalog,
} from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { insertTable } from './ribbon-commands';
import { tableCellStyle } from './table-render';

registerDocxEditor();

const cell = (id: string, text = id) => ({
	paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text }] }],
});
function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
const cells = (editor: DocxEditorElement) => [
	...editor.shadowRoot!.querySelectorAll<HTMLTableCellElement>('.dve-paper td'),
];

describe('table rendering', () => {
	afterEach(() => document.body.replaceChildren());

	it('renders table style banding, header shading and inside borders', () => {
		const styles: TableStyleCatalog = {
			styles: {
				Grid: {
					id: 'Grid',
					borders: {
						top: { style: 'single', sizeEighthPoints: 8, color: '#000000' },
						bottom: { style: 'single', sizeEighthPoints: 8, color: '#000000' },
						insideH: { style: 'dashed', sizeEighthPoints: 4, color: '#999999' },
					},
					conditional: { firstRow: { shadingFill: '#1f4e79' } },
				},
			},
			warnings: [],
		};
		const model = createDocument();
		model.tableStyles = styles;
		model.blocks = [
			{
				type: 'table',
				id: 't1',
				style: 'Grid',
				look: { firstRow: true, noHBand: true, noVBand: true },
				rows: [
					[cell('a'), cell('b')],
					[cell('c'), cell('d')],
				],
			},
		];
		const [a, , c] = cells(mount(model));
		expect(a.style.backgroundColor).toBe('rgb(31, 78, 121)');
		expect(c.style.backgroundColor).toBe('');
		expect(a.style.borderTop).toContain('solid');
		expect(a.style.borderBottom).toContain('dashed');
		// jsdom drops border "none" declarations; check the generated CSS instead.
		expect(
			tableCellStyle({
				borders: JSON.stringify({ left: { style: 'none' }, top: { style: 'single' } }),
			}),
		).toContain('border-left-style:none');
	});

	it('shows dashed gridlines only for tables without any border information', () => {
		const model = createDocument();
		model.blocks = [{ type: 'table', id: 't1', rows: [[cell('a')]] }];
		const [only] = cells(mount(model));
		expect(only.getAttribute('style') ?? '').not.toContain('border-top');
	});

	it('inserts tables with Word default borders and saves a valid tblPr and tblGrid', async () => {
		const editor = mount(createDocument());
		const view = (editor as unknown as { view: EditorView }).view;
		insertTable(view);
		const table = editor.documentModel!.blocks.find((block) => block.type === 'table') as Table;
		expect(table.borders?.insideH).toMatchObject({ style: 'single', sizeEighthPoints: 4 });
		const saved = await saveDocx(structuredClone(editor.documentModel!));
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toMatch(/<w:tbl><w:tblPr><w:tblW [^>]*\/><w:tblBorders>/);
		expect(xml.match(/<w:gridCol /g)).toHaveLength(2);
		const reloaded = await loadDocx(saved);
		const reloadedTable = reloaded.model.blocks.find((block) => block.type === 'table') as Table;
		expect(reloadedTable.borders?.top).toMatchObject({ style: 'single', sizeEighthPoints: 4 });
		expect(reloadedTable.grid).toHaveLength(2);
	});
});
