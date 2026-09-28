import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index.js';
import { at, expectTable } from './test-support/access.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (text: string, fill: string, width: number) =>
	`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="${fill}"/></w:tcPr><w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const sourceXml = `<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr/><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:shd w:fill="FF0001"/></w:tcPr><w:p><w:r><w:rPr><w:lang w:val="en-US"/></w:rPr><w:t>A</w:t></w:r></w:p></w:tc>${cell('B', 'FF0002', 2000)}</w:tr><w:tr>${cell('C', 'FF0003', 1000)}${cell('D', 'FF0004', 2000)}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`;

async function loadFixture(xml = sourceXml) {
	const zip = new JSZip();
	zip.file('word/document.xml', xml);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const table = expectTable(loaded.model.blocks[0]);
	return { loaded, table };
}

function emptyCell(id: string) {
	return { paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text: '' }] }] };
}

async function savedTableXml(loaded: Awaited<ReturnType<typeof loadDocx>>) {
	const zip = await JSZip.loadAsync(await loaded.save());
	return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

function tableCells(xml: string): string[] {
	const table = xml.match(/<w:tbl(?:\s[^>]*)?>[\s\S]*?<\/w:tbl>/)?.[0] ?? '';
	return [...table.matchAll(/<w:tc(?:\s[^>]*)?>([\s\S]*?)<\/w:tc>/g)].map((match) => at(match, 1));
}

function assertCellIdentity(cells: string[], text: string, fill: string, width: number) {
	const found = cells.find((candidate) => candidate.includes(`<w:t>${text}</w:t>`));
	expect(found, `cell ${text} should retain its XML`).toBeDefined();
	expect(found).toContain(`w:fill="${fill}"`);
	expect(found).toContain(`w:w="${width}"`);
	expect(found).toContain('<w:lang w:val="en-US"');
}

describe('DOCX table structure editing', () => {
	it('inserts a row while keeping each original cell XML with its paragraph identity', async () => {
		const { loaded, table } = await loadFixture();
		table.rows.splice(1, 0, [emptyCell('new-row-left'), emptyCell('new-row-right')]);
		const xml = await savedTableXml(loaded);
		const cells = tableCells(xml);
		expect(cells).toHaveLength(6);
		assertCellIdentity(cells, 'A', 'FF0001', 1000);
		assertCellIdentity(cells, 'B', 'FF0002', 2000);
		assertCellIdentity(cells, 'C', 'FF0003', 1000);
		assertCellIdentity(cells, 'D', 'FF0004', 2000);
		expect(xml.match(/<w:gridCol\b/g)).toHaveLength(2);
	});

	it('deletes a row and retains the remaining row cell properties', async () => {
		const { loaded, table } = await loadFixture();
		table.rows.splice(0, 1);
		const xml = await savedTableXml(loaded);
		const cells = tableCells(xml);
		expect(cells).toHaveLength(2);
		assertCellIdentity(cells, 'C', 'FF0003', 1000);
		assertCellIdentity(cells, 'D', 'FF0004', 2000);
		expect(xml.match(/<w:gridCol\b/g)).toHaveLength(2);
	});

	it('inserts a column consistently across rows and updates tblGrid', async () => {
		const { loaded, table } = await loadFixture();
		for (const [index, row] of table.rows.entries())
			row.splice(1, 0, emptyCell(`new-column-${index}`));
		const xml = await savedTableXml(loaded);
		const cells = tableCells(xml);
		expect(cells).toHaveLength(6);
		assertCellIdentity(cells, 'A', 'FF0001', 1000);
		assertCellIdentity(cells, 'B', 'FF0002', 2000);
		assertCellIdentity(cells, 'C', 'FF0003', 1000);
		assertCellIdentity(cells, 'D', 'FF0004', 2000);
		expect(xml.match(/<w:gridCol\b/g)).toHaveLength(3);
	});

	it('deletes a column while keeping the surviving column cell XML and grid entry', async () => {
		const { loaded, table } = await loadFixture();
		for (const row of table.rows) row.splice(0, 1);
		const xml = await savedTableXml(loaded);
		const cells = tableCells(xml);
		expect(cells).toHaveLength(2);
		assertCellIdentity(cells, 'B', 'FF0002', 2000);
		assertCellIdentity(cells, 'D', 'FF0004', 2000);
		expect(xml.match(/<w:gridCol\b/g)).toHaveLength(1);
		expect(xml).toContain('<w:gridCol w:w="2000"');
	});

	it.each([
		['inconsistent grid', sourceXml.replace('<w:gridCol w:w="2000"/>', '')],
		['merged cells', sourceXml.replace('<w:tcPr>', '<w:tcPr><w:gridSpan w:val="2"/>')],
		[
			'nested tables',
			sourceXml.replace(
				'</w:p></w:tc>',
				'</w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc>',
			),
		],
		[
			'hidden inline content',
			sourceXml.replace('<w:t>A</w:t>', '<w:bookmarkStart w:id="1" w:name="hidden"/><w:t>A</w:t>'),
		],
	])(
		'rejects structural edits to %s tables while preserving no-op package bytes',
		async (_name, xml) => {
			const zip = new JSZip();
			zip.file('word/document.xml', xml);
			const original = await zip.generateAsync({ type: 'uint8array' });
			const loaded = await loadDocx(original);
			const table = expectTable(loaded.model.blocks[0]);
			expect(table.structureEditable).toBe(false);
			expect(loaded.model.warnings).toContain(
				'Merged, nested, or complex tables can be read, but their row and column structure cannot be edited safely.',
			);
			expect(await loaded.save()).toEqual(original);
			table.rows.splice(1, 0, [emptyCell('new-left'), emptyCell('new-right')]);
			await expect(loaded.save()).rejects.toThrow('Cannot change the structure');
			expect(await (await loadDocx(original)).save()).toEqual(original);
		},
	);
});
