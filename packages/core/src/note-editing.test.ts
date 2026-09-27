import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index.js';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}"><w:footnote w:id="0" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Source: survey.</w:t></w:r></w:p></w:footnote></w:footnotes>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('footnote editing', () => {
	it('models the note number mark and rewrites only the edited note', async () => {
		const loaded = await loadDocx(await fixture());
		const note = loaded.model.footnotes![0];
		const paragraph = note.blocks[0] as Paragraph;
		expect(paragraph.runs[0]).toMatchObject({ text: '', noteMark: 'footnote' });
		paragraph.runs[1].text = ' Source: 2026 survey.';
		const saved = await loaded.save(loaded.model);
		const xml = await (await JSZip.loadAsync(saved)).file('word/footnotes.xml')!.async('string');
		expect(xml).toContain('2026 survey.');
		expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/><\/w:rPr><w:footnoteRef\/>/);
		expect(xml).toContain('<w:separator/>');
		expect(xml).toContain('<w:continuationSeparator/>');
		const reloaded = await loadDocx(saved);
		expect((reloaded.model.footnotes![0].blocks[0] as Paragraph).runs[1].text).toBe(
			' Source: 2026 survey.',
		);
	});

	it('adds a new note to an existing footnotes part', async () => {
		const loaded = await loadDocx(await fixture());
		loaded.model.footnotes!.push({
			id: '3',
			blocks: [
				{
					type: 'paragraph',
					id: 'fn3-p0',
					runs: [
						{ text: '', noteMark: 'footnote', verticalAlign: 'superscript' },
						{ text: ' New note' },
					],
				},
			],
		});
		(loaded.model.blocks[0] as Paragraph).runs.push({
			text: '',
			noteReference: { kind: 'footnote', id: '3' },
			verticalAlign: 'superscript',
		});
		const saved = await loaded.save(loaded.model);
		const xml = await (await JSZip.loadAsync(saved)).file('word/footnotes.xml')!.async('string');
		expect(xml).toMatch(/<w:footnote w:id="3"><w:p>.*<w:footnoteRef\/>.*New note/);
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.footnotes!.map((note) => note.id)).toEqual(['2', '3']);
	});

	it('creates the footnotes part, relationship and content type for a first footnote', async () => {
		const zip = new JSZip();
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
		);
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		loaded.model.footnotes = [
			{
				id: '1',
				blocks: [
					{
						type: 'paragraph',
						id: 'fn1-p0',
						runs: [{ text: '', noteMark: 'footnote' }, { text: ' First' }],
					},
				],
			},
		];
		(loaded.model.blocks[0] as Paragraph).runs.push({
			text: '',
			noteReference: { kind: 'footnote', id: '1' },
		});
		const saved = await JSZip.loadAsync(await loaded.save(loaded.model));
		const notes = await saved.file('word/footnotes.xml')!.async('string');
		expect(notes).toContain('w:type="separator" w:id="-1"');
		expect(notes).toContain(' First');
		expect(await saved.file('word/_rels/document.xml.rels')!.async('string')).toContain(
			'Target="footnotes.xml"',
		);
		expect(await saved.file('[Content_Types].xml')!.async('string')).toContain(
			'PartName="/word/footnotes.xml"',
		);
		const reloaded = await loadDocx(await saved.generateAsync({ type: 'uint8array' }));
		expect(reloaded.model.footnotes).toHaveLength(1);
	});
});
