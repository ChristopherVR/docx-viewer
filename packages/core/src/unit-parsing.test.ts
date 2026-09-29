import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type Paragraph, type Table } from './index.js';
import { at, must } from './test-support/access.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
async function load(body: string) {
	const zip = new JSZip();
	zip.file('word/document.xml', `<w:document ${NS}><w:body>${body}</w:body></w:document>`);
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

describe('signed and unsigned twip fields follow the schema when parsing', () => {
	it('keeps signed indents and tab positions but drops negative unsigned values', async () => {
		const model = await load(
			'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="-360"/></w:tabs>' +
				'<w:spacing w:before="-20" w:after="120" w:line="-240"/>' +
				'<w:ind w:left="-720" w:hanging="-360" w:firstLine="180"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p><w:sectPr/>',
		);
		const paragraph = at(model.blocks, 0) as Paragraph;
		expect(paragraph.tabStops?.[0]?.posTwips).toBe(-360);
		expect(paragraph.indentLeftTwips).toBe(-720);
		expect(paragraph.lineSpacingTwips).toBe(-240);
		expect(paragraph.spacingAfterTwips).toBe(120);
		expect(paragraph.firstLineTwips).toBe(180);
		expect(paragraph.spacingBeforeTwips).toBeUndefined();
		expect(paragraph.hangingTwips).toBeUndefined();
	});

	it('keeps a negative table indent, the usual Word "align text with the margin" value', async () => {
		const model = await load(
			'<w:tbl><w:tblPr><w:tblInd w:w="-108" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
				'<w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr/>',
		);
		const table = at(model.blocks, 0) as Table;
		expect(table.indentTwips).toBe(-108);
		expect(table.grid).toEqual([4000]);
	});

	it('accepts negative top/bottom page margins but not negative side margins', async () => {
		const model = await load(
			'<w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="-200" w:right="-1" w:bottom="-400" w:left="720"/></w:sectPr>',
		);
		const section = must(model.sections?.at(-1), 'final section');
		expect(section.marginTopTwips).toBe(-200);
		expect(section.marginBottomTwips).toBe(-400);
		expect(section.marginRightTwips).toBe(1440);
		expect(section.marginLeftTwips).toBe(720);
	});
});
