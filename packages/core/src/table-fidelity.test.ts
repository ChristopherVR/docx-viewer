import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, parseTableStyleCatalog, resolveTableStyleFormatting } from './index.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const mergedTableXml = `<w:document xmlns:w="${ns}"><w:body>
<w:tbl>
<w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="5000" w:type="dxa"/><w:jc w:val="center"/><w:tblInd w:w="100"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="FF0000"/></w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:noHBand="0"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/><w:shd w:fill="ABCDEF"/><w:tcMar><w:top w:w="50" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="1"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>D</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr>
</w:tbl>
<w:sectPr/></w:body></w:document>`;

async function loadTable(xml = mergedTableXml, styles?: string) {
	const zip = new JSZip();
	zip.file('word/document.xml', xml);
	if (styles) zip.file('word/styles.xml', styles);
	const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const table = loaded.model.blocks[0];
	if (table.type !== 'table') throw new Error('Expected a table fixture');
	return { loaded, table };
}

describe('table grid, merge, border, shading and style fidelity', () => {
	it('parses grid widths, table width/alignment/indent, borders and tblLook', async () => {
		const { table } = await loadTable();
		expect(table.grid).toEqual([1000, 2000]);
		expect(table.widthTwips).toBe(5000);
		expect(table.alignment).toBe('center');
		expect(table.indentTwips).toBe(100);
		expect(table.style).toBe('Grid');
		expect(table.borders?.top).toMatchObject({ style: 'single', sizeEighthPoints: 8, color: '#FF0000' });
		expect(table.look).toMatchObject({ firstRow: true, noHBand: false });
	});

	it('parses gridSpan, vMerge restart/continue, cell width, vAlign, shading and margins', async () => {
		const { table } = await loadTable();
		const [[a, b], [c, d]] = table.rows;
		expect(a).toMatchObject({
			widthTwips: 1000,
			verticalMerge: 'restart',
			verticalAlign: 'center',
			shadingFill: '#ABCDEF',
			margins: { top: 50 },
		});
		expect(b.gridSpan).toBe(1);
		expect(c.verticalMerge).toBe('continue');
		expect(d.widthTwips).toBe(2000);
	});

	it('surfaces a read-only text preview for a nested table instead of dropping it', async () => {
		const { table, loaded } = await loadTable();
		const nested = table.rows[1][1].nestedTables;
		expect(nested).toHaveLength(1);
		expect(nested?.[0].rows).toEqual([[{ text: 'Nested' }]]);
		expect(loaded.model.warnings.join(' ')).toContain('Nested tables render as a read-only text preview');
	});

	it('marks the table with a nested tbl as not structurally editable, and a no-op save keeps original bytes', async () => {
		const zip = new JSZip();
		zip.file('word/document.xml', mergedTableXml);
		const original = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(original);
		const table = loaded.model.blocks[0];
		if (table.type !== 'table') throw new Error('Expected a table fixture');
		expect(table.structureEditable).toBe(false);
		expect(await loaded.save()).toEqual(original);
	});

	it('rejects silently dropping a direct edit to table-level descriptive properties', async () => {
		const { table, loaded } = await loadTable();
		table.widthTwips = 9999;
		await expect(loaded.save()).rejects.toThrow('Cannot edit table grid widths, width, alignment');
	});

	it('rejects silently dropping a direct edit to cell-level descriptive properties', async () => {
		const { table, loaded } = await loadTable();
		table.rows[0][0].shadingFill = '#000000';
		await expect(loaded.save()).rejects.toThrow('Cannot edit table cell width, merge');
	});

	it('parses table style conditional formatting (tblStylePr) and resolves banding/first-row precedence', () => {
		const stylesXml = `<w:styles xmlns:w="${ns}"><w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/>
<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:fill="4472C4"/></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
</w:style></w:styles>`;
		const catalog = parseTableStyleCatalog(stylesXml);
		expect(catalog.styles.Grid.conditional.firstRow?.shadingFill).toBe('#4472C4');
		expect(catalog.styles.Grid.conditional.firstRow?.run).toMatchObject({ bold: true });
		expect(catalog.styles.Grid.borders?.top).toMatchObject({ style: 'single' });
		// Row 0 is both banded (band1Horz) and firstRow; firstRow has higher precedence.
		const firstRow = resolveTableStyleFormatting('Grid', catalog, { firstRow: true }, 0, 3, 0, 2);
		expect(firstRow.shadingFill).toBe('#4472C4');
		// Row 2 (not first/last, even index => band1Horz) falls back to banding only.
		const bandedRow = resolveTableStyleFormatting('Grid', catalog, { firstRow: true }, 2, 3, 0, 2);
		expect(bandedRow.shadingFill).toBe('#D9E2F3');
	});

	it('parses styles.xml for the table catalog when loading a package and reports the fidelity warning', async () => {
		const { loaded } = await loadTable(
			mergedTableXml,
			`<w:styles xmlns:w="${ns}"><w:style w:type="table" w:styleId="Grid"><w:tblPr/></w:style></w:styles>`,
		);
		expect(loaded.model.tableStyles?.styles.Grid).toBeDefined();
		expect(loaded.model.warnings.join(' ')).toContain('table style conditional formatting resolve for rendering');
	});
});
