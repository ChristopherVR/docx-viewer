import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Table } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('table row properties and default cell margins', () => {
	it('parses w:tblCellMar and w:trPr, and keeps them when cell text is edited', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:start w:w="0" w:type="dxa"/><w:end w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>` +
				`<w:tr><w:trPr><w:tblHeader/><w:cantSplit/><w:trHeight w:val="600" w:hRule="exact"/></w:trPr><w:tc><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:tc></w:tr>` +
				`<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr><w:tc><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:tc></w:tr>` +
				`<w:tr><w:tc><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const table = loaded.model.blocks[0] as Table;
		expect(table.cellMargins).toEqual({ top: 60, left: 0, right: 0 });
		expect(table.rowProperties).toEqual([
			{ heightTwips: 600, heightRule: 'exact', cantSplit: true, header: true },
			{ heightTwips: 400, heightRule: 'atLeast' },
			{},
		]);
		table.rows[1][0].paragraphs[0].runs[0].text = 'Edited';
		const xml = await (
			await JSZip.loadAsync(await loaded.save(loaded.model))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('<w:trHeight w:val="600" w:hRule="exact"/>');
		expect(xml).toContain('<w:tblHeader/>');
		expect(xml).toContain('Edited');
	});
});
