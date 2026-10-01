import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { ensureListDefinition } from './numbering-editing.js';
import { computeListLabels } from './numbering-format.js';
import { twips } from './units.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/numbering.xml',
		`<?xml version="1.0"?><w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('numbering marker format', () => {
	it('parses level rPr onto the level and the computed label', async () => {
		const { model } = await loadDocx(await fixture());
		const expected = {
			fontFamily: 'Georgia',
			bold: true,
			color: '#1f4e79',
			fontSizeHalfPoints: 28,
		};
		expect(model.numberingCatalog!.abstractNums['0']!.levels[0]!.markerFormat).toEqual(expected);
		expect([...computeListLabels(model).values()][0]!.markerFormat).toEqual(expected);
	});

	it('writes marker formatting in schema order for a new definition and reads it back', async () => {
		const loaded = await loadDocx(await fixture());
		const added = ensureListDefinition(loaded.model.numberingCatalog, 'decimal');
		const abstract = added.catalog.abstractNums[added.catalog.nums[added.numId]!.abstractNumId]!;
		abstract.levels[0]!.markerFormat = { fontFamily: 'Arial', italic: true, color: '#ff0000' };
		loaded.model.numberingCatalog = added.catalog;
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/numbering.xml')?.async('string')) ?? '';
		expect(xml).toMatch(
			/<w:rPr><w:rFonts [^>]*w:ascii="Arial"[^>]*\/><w:i\/><w:color w:val="FF0000"\/><\/w:rPr>/,
		);
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		const id = added.catalog.nums[added.numId]!.abstractNumId;
		expect(reopened.model.numberingCatalog!.abstractNums[id]!.levels[0]!.markerFormat).toEqual({
			fontFamily: 'Arial',
			italic: true,
			color: '#ff0000',
		});
	});

	it('round-trips a level tab stop before the indent', async () => {
		const loaded = await loadDocx(await fixture());
		const added = ensureListDefinition(loaded.model.numberingCatalog, 'decimal');
		const id = added.catalog.nums[added.numId]!.abstractNumId;
		added.catalog.abstractNums[id]!.levels[0]!.tabStopTwips = twips(1440);
		loaded.model.numberingCatalog = added.catalog;
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/numbering.xml')?.async('string')) ?? '';
		expect(xml).toMatch(/<w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"\/><\/w:tabs><w:ind /);
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		expect(reopened.model.numberingCatalog!.abstractNums[id]!.levels[0]!.tabStopTwips).toBe(1440);
	});
});
