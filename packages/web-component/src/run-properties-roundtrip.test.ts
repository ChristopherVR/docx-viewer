import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph, type Table } from 'docx-core';
import { docToModel, modelToDoc } from './model-adapter';
import { at, paragraphAt, tableAt } from './test-support';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function docx(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Converts a model into the editor document and back, as every editor transaction does. */
const throughEditor = <T extends Parameters<typeof modelToDoc>[0]>(model: T) =>
	docToModel(modelToDoc(model), model);

describe('run properties survive the editor round trip', () => {
	it('keeps character style, caps and theme colour when text around them is edited', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:p><w:r><w:rPr><w:rStyle w:val="Emph"/><w:caps/><w:color w:val="1F3864" w:themeColor="accent1" w:themeShade="BF"/></w:rPr><w:t>Styled</w:t></w:r><w:r><w:t> plain</w:t></w:r></w:p>',
			),
		);
		const roundTripped = throughEditor(loaded.model);
		const edited = structuredClone(roundTripped);
		at(paragraphAt(edited.blocks, 0).runs, 1).text = ' edited';
		const styled = at(paragraphAt(edited.blocks, 0).runs, 0);
		expect(styled).toMatchObject({ style: 'Emph', caps: true, colorTheme: { token: 'accent1' } });
		const saved = await loaded.save(Object.assign(loaded.model, { blocks: edited.blocks }));
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain('w:rStyle w:val="Emph"');
		expect(xml).toContain('<w:caps');
		expect(xml).toContain('w:themeColor="accent1"');
		expect(xml).toContain(' edited');
	});

	it('keeps table grid, borders and style after an editor round trip of a simple table', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
			),
		);
		const roundTripped = throughEditor(loaded.model);
		const table = roundTripped.blocks[0] as Table;
		expect(table.style).toBe('Grid');
		expect(table.grid).toEqual([2000]);
		expect(table.borders).toEqual((loaded.model.blocks[0] as Table).borders);
	});

	it('saves text edits in a simple table whose cells carry width and shading', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
			),
		);
		const edited = throughEditor(loaded.model);
		const cell = at(at(tableAt(edited.blocks, 0).rows, 0), 0);
		expect(cell).toMatchObject({ widthTwips: 2000, shadingFill: expect.stringMatching(/ffff00/i) });
		at(at(cell.paragraphs, 0).runs, 0).text = 'Edited cell';
		const saved = await loaded.save(Object.assign(loaded.model, { blocks: edited.blocks }));
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain('Edited cell');
		expect(xml).toContain('w:fill="FFFF00"');
	});

	it('keeps a footnote reference run formatting through the editor', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:p><w:r><w:t>See</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p>',
			),
		);
		const edited = throughEditor(loaded.model);
		expect((edited.blocks[0] as Paragraph).runs[1]).toMatchObject({
			noteReference: { kind: 'footnote', id: '2' },
			style: 'FootnoteReference',
			verticalAlign: 'superscript',
		});
	});
});
