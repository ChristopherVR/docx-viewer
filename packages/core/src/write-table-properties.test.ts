import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, signedTwips, type Table } from './index.js';

describe('existing table property edits', () => {
	it('patches modeled row and margin fields, keeping unrelated source XML', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>
<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="10" w:type="dxa"/><w:start w:w="80" w:type="dxa"/></w:tblCellMar><w:tblDescription w:val="Keep me"/></w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="400" w:hRule="exact"/><w:tblHeader/><w:hidden/></w:trPr><w:tc><w:tcPr><w:noWrap/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const model = structuredClone(loaded.model);
		const table = model.blocks[0] as Table;
		table.rowProperties = [{}];
		table.cellMargins = { ...table.cellMargins, top: signedTwips(144) };
		const saved = await loaded.save(model);
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).not.toMatch(/<w:(cantSplit|trHeight|tblHeader)/);
		for (const preserved of [
			'<w:hidden',
			'<w:noWrap',
			'<w:tblLayout w:type="fixed"',
			'<w:start w:w="80"',
			'<w:tblDescription w:val="Keep me"',
		])
			expect(xml).toContain(preserved);
		expect(xml).toContain('<w:top w:w="144" w:type="dxa"');
		const reopened = await loadDocx(saved);
		expect((reopened.model.blocks[0] as Table).rowProperties).toBeUndefined();
		expect((reopened.model.blocks[0] as Table).cellMargins?.top).toBe(144);
	});
});
