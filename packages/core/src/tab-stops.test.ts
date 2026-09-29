import { signedTwips, twips } from './units.js';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx, type Paragraph } from './index.js';

async function documentXml(bytes: Uint8Array) {
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

describe('paragraph properties', () => {
	it('writes new paragraph properties in schema order', async () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [{ text: 'Title' }],
				align: 'center',
				style: 'Heading1',
				pageBreakBefore: true,
				spacingAfterTwips: twips(120),
				tabStops: [{ posTwips: signedTwips(4680), align: 'center' }],
			},
		];
		expect(await documentXml(await saveDocx(model))).toContain(
			'<w:pPr><w:pStyle w:val="Heading1"/><w:pageBreakBefore/><w:tabs><w:tab w:val="center" w:pos="4680"/></w:tabs><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr>',
		);
	});

	it('keeps and edits tab stops of loaded paragraphs', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/><w:tab w:val="bogus" w:pos="1"/></w:tabs><w:jc w:val="left"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.tabStops).toEqual([{ posTwips: 9000, align: 'right', leader: 'dot' }]);
		const edited = { ...paragraph, runs: [{ text: 'B' }], align: 'right' as const };
		const xml = await documentXml(await loaded.save({ ...loaded.model, blocks: [edited] }));
		expect(xml).toContain(
			'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/><w:tab w:val="bogus" w:pos="1"/></w:tabs><w:jc w:val="right"/>',
		);
		const moved = {
			...paragraph,
			tabStops: [{ posTwips: signedTwips(500), align: 'left' as const }],
		};
		expect(await documentXml(await loaded.save({ ...loaded.model, blocks: [moved] }))).toContain(
			'<w:tabs><w:tab w:val="left" w:pos="500"/></w:tabs>',
		);
	});
});
