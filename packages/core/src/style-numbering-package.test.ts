import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { createListDefinition, linkStylesToList } from './numbering-editing.js';
import { computeListLabels } from './numbering-format.js';
import type { Paragraph } from './model.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Sub</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/styles.xml',
		`<?xml version="1.0"?><w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr><w:i/></w:rPr></w:style></w:styles>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const levels = [
	{ level: 0, start: 1, numFmt: 'decimal' as const, lvlText: '%1.', paragraphStyleId: 'Heading1' },
	{
		level: 1,
		start: 1,
		numFmt: 'decimal' as const,
		lvlText: '%1.%2.',
		paragraphStyleId: 'Heading2',
	},
];

describe('heading-linked list styles', () => {
	it('writes style numPr in schema order, keeps other style XML, and numbers the headings', async () => {
		const loaded = await loadDocx(await fixture());
		const created = createListDefinition(loaded.model.numberingCatalog, levels);
		loaded.model.numberingCatalog = created.catalog;
		loaded.model.paragraphStyles = linkStylesToList(
			loaded.model.paragraphStyles,
			levels,
			created.numId,
		)!;
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/styles.xml')?.async('string')) ?? '';
		expect(xml).toMatch(
			/styleId="Heading1">.*<w:pPr><w:keepNext\/><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="1"\/><\/w:numPr><w:spacing w:before="240"\/><\/w:pPr><w:rPr><w:b\/>/,
		);
		expect(xml).toMatch(
			/styleId="Heading2">.*<w:pPr><w:numPr><w:ilvl w:val="1"\/><w:numId w:val="1"\/><\/w:numPr><\/w:pPr><w:rPr><w:i\/>/,
		);
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		const labels = computeListLabels(reopened.model);
		const ids = reopened.model.blocks.map((block) => (block as Paragraph).id);
		expect(ids.map((id) => labels.get(id)?.text)).toEqual(['1.', '1.1.']);
	});

	it('still rejects other style catalog edits', async () => {
		const loaded = await loadDocx(await fixture());
		loaded.model.paragraphStyles!.styles.Heading1!.name = 'Renamed';
		await expect(loaded.save()).rejects.toThrow(/Editing the paragraph style catalog/);
	});

	it('removes a link when numbering is cleared', async () => {
		const loaded = await loadDocx(await fixture());
		const created = createListDefinition(loaded.model.numberingCatalog, levels);
		loaded.model.numberingCatalog = created.catalog;
		loaded.model.paragraphStyles = linkStylesToList(loaded.model.paragraphStyles, levels, 1)!;
		const linked = await loadDocx(await loaded.save());
		delete linked.model.paragraphStyles!.styles.Heading1!.numbering;
		const xml =
			(await (
				await JSZip.loadAsync(await linked.save())
			)
				.file('word/styles.xml')
				?.async('string')) ?? '';
		expect(xml).not.toMatch(/Heading1">.*<w:numPr>.*Heading2"/);
		expect(xml).toMatch(/Heading2">.*<w:numPr>/);
	});
});
