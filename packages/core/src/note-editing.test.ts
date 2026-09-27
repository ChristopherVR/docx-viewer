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

	it('rejects notes without a matching note element', async () => {
		const loaded = await loadDocx(await fixture());
		loaded.model.footnotes!.push({
			id: '9',
			blocks: [{ type: 'paragraph', id: 'fn9-p0', runs: [{ text: 'New' }] }],
		});
		await expect(loaded.save(loaded.model)).rejects.toThrow(/not supported yet/);
	});
});
