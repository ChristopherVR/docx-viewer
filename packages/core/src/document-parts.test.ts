import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { formatNoteNumber, loadDocx, numberNotesInOrder, saveDocx } from './index.js';
import { at, expectParagraph } from './test-support/access.js';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function fixtureZip(): Promise<JSZip> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body>
<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
<w:p><w:r><w:t>Unrelated</w:t></w:r></w:p>
<w:sectPr>
<w:headerReference w:type="default" r:id="rId1"/>
<w:footerReference w:type="default" r:id="rId2"/>
<w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
</w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
	);
	zip.file(
		'word/footer1.xml',
		`<w:ftr xmlns:w="${w}"><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
	);
	zip.file(
		'word/settings.xml',
		`<w:settings xmlns:w="${w}"><w:evenAndOddHeaders/><w:footnotePr><w:numFmt w:val="decimal"/></w:footnotePr><w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr></w:settings>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}">
<w:footnote w:id="0" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:id="1" w:type="continuationSeparator"><w:p><w:r><w:t/></w:r></w:p></w:footnote>
<w:footnote w:id="2"><w:p><w:r><w:t>Footnote body</w:t></w:r></w:p></w:footnote>
</w:footnotes>`,
	);
	return zip;
}

describe('headers, footers and settings', () => {
	it('resolves default header/footer content per section and reports evenAndOddHeaders', async () => {
		const zip = await fixtureZip();
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		const section = loaded.model.sections?.[0];
		expect(section?.headers?.default?.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: 'Header text' }],
		});
		expect(section?.footers?.default?.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: '1', field: { instr: 'PAGE' } }],
		});
		expect(loaded.model.evenAndOddHeaders).toBe(true);
		expect(loaded.model.warnings.some((warning) => /cannot be edited/.test(warning))).toBe(false);
		expect(loaded.model.warnings).toContain(
			'PAGE, NUMPAGES, SECTIONPAGES, DATE and TIME fields are recalculated in Print Layout; other fields show the result Word last saved. Field results are editable and field codes are preserved.',
		);
	});

	it('preserves header/footer/notes parts byte-exact on an unrelated body edit', async () => {
		const zip = await fixtureZip();
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[1]);
		paragraph.align = 'center';
		const saved = await loaded.save();
		const savedZip = await JSZip.loadAsync(saved);
		expect(await savedZip.file('word/header1.xml')?.async('string')).toEqual(
			await zip.file('word/header1.xml')?.async('string'),
		);
		expect(await savedZip.file('word/footnotes.xml')?.async('string')).toEqual(
			await zip.file('word/footnotes.xml')?.async('string'),
		);
	});

	it('rejects a model whose sections were edited instead of silently dropping the change', async () => {
		const zip = await fixtureZip();
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		loaded.model.sections = [];
		await expect(loaded.save()).rejects.toThrow('at least one section');
	});
});

describe('footnotes', () => {
	it('parses footnote bodies while skipping Word separator marks', async () => {
		const zip = await fixtureZip();
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		expect(loaded.model.footnotes).toHaveLength(1);
		expect(loaded.model.footnotes?.[0]).toMatchObject({
			id: '2',
			blocks: [{ type: 'paragraph', runs: [{ text: 'Footnote body' }] }],
		});
		expect(loaded.model.footnoteNumFmt).toBe('decimal');
		expect(loaded.model.endnoteNumFmt).toBe('lowerRoman');
	});

	it('numbers notes by first-reference order and formats per numFmt', () => {
		const order = numberNotesInOrder(
			[
				{
					type: 'paragraph',
					id: 'p0',
					runs: [{ text: '' }, { text: '', noteReference: { kind: 'footnote', id: '7' } }],
				},
				{
					type: 'paragraph',
					id: 'p1',
					runs: [{ text: '', noteReference: { kind: 'footnote', id: '9' } }],
				},
			],
			'footnote',
		);
		expect(order.get('7')).toBe(1);
		expect(order.get('9')).toBe(2);
		expect(formatNoteNumber(4, 'upperRoman')).toBe('IV');
		expect(formatNoteNumber(2, 'lowerLetter')).toBe('b');
		expect(formatNoteNumber(3, undefined)).toBe('3');
	});
});

describe('saveDocx guards', () => {
	it('writes footnotes from a standalone (non-loaded) model into a new footnotes part', async () => {
		const saved = await saveDocx({
			blocks: [
				{
					type: 'paragraph',
					id: 'p1',
					runs: [{ text: 'Text' }, { text: '', noteReference: { kind: 'footnote', id: '1' } }],
				},
			],
			page: {
				width: 816,
				height: 1056,
				marginTop: 96,
				marginRight: 96,
				marginBottom: 96,
				marginLeft: 96,
			},
			warnings: [],
			footnotes: [
				{
					id: '1',
					blocks: [
						{
							type: 'paragraph',
							id: 'fn1-p0',
							runs: [{ text: '', noteMark: 'footnote' }, { text: ' Note' }],
						},
					],
				},
			],
		});
		const reloaded = await loadDocx(saved);
		expect(at(reloaded.model.footnotes, 0).blocks[0]).toMatchObject({
			runs: [{ noteMark: 'footnote' }, { text: ' Note' }],
		});
	});
});
