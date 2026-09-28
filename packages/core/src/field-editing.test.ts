import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index.js';
import { at } from './test-support/access.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function docx(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function documentXml(bytes: Uint8Array): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

const complexField = `<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Figure 1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> above.</w:t></w:r></w:p>`;

describe('editable fields', () => {
	it('models complex field markers and codes as runs', async () => {
		const loaded = await loadDocx(await docx(complexField));
		const runs = (loaded.model.blocks[0] as Paragraph).runs;
		expect(runs).toEqual([
			{ text: 'See ' },
			{ text: '', fieldChar: 'begin', bold: true },
			{ text: '', fieldCode: ' REF _Ref1 \\h ' },
			{ text: '', fieldChar: 'separate' },
			{ text: 'Figure 1', field: { instr: 'REF _Ref1 \\h' } },
			{ text: '', fieldChar: 'end' },
			{ text: ' above.' },
		]);
	});

	it('saves edits around and inside a complex field, keeping its structure', async () => {
		const loaded = await loadDocx(await docx(complexField));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const runs = paragraph.runs.map((run) => ({ ...run }));
		at(runs, 0).text = 'Compare ';
		at(runs, 4).text = 'Figure 2';
		at(runs, 6).text = ' below.';
		const next = { ...loaded.model, blocks: [{ ...paragraph, runs }] };
		const xml = await documentXml(await loaded.save(next));
		expect(xml).toMatch(
			/Compare <\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t>Figure 2<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/><\/w:r><w:r><w:t xml:space="preserve"> below.<\/w:t>/,
		);
		const reloaded = await loadDocx(await loaded.save(next));
		expect((reloaded.model.blocks[0] as Paragraph).runs[4]).toEqual({
			text: 'Figure 2',
			field: { instr: 'REF _Ref1 \\h' },
		});
	});

	it('edits simple field results and keeps them in w:fldSimple', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:r><w:t xml:space="preserve">By </w:t></w:r><w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ann</w:t></w:r></w:fldSimple></w:p>`,
			),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.runs[1]).toEqual({ text: 'Ann', field: { instr: 'AUTHOR', simple: true } });
		const runs = [{ text: 'Written by ' }, { ...paragraph.runs[1], text: 'Bea' }];
		const xml = await documentXml(
			await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] }),
		);
		expect(xml).toContain('<w:t xml:space="preserve">Written by </w:t>');
		expect(xml).toMatch(
			/<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Bea<\/w:t><\/w:r><\/w:fldSimple>/,
		);
	});

	it('writes new complex fields from marker runs', async () => {
		const loaded = await loadDocx(await docx('<w:p><w:r><w:t>Intro</w:t></w:r></w:p>'));
		const paragraph = loaded.model.blocks[0] as Paragraph;
		const runs = [
			...paragraph.runs,
			{ text: '', fieldChar: 'begin' as const },
			{ text: '', fieldCode: ' PAGE ' },
			{ text: '', fieldChar: 'separate' as const },
			{ text: '1', field: { instr: 'PAGE' } },
			{ text: '', fieldChar: 'end' as const },
		];
		const saved = await loaded.save({ ...loaded.model, blocks: [{ ...paragraph, runs }] });
		expect(await documentXml(saved)).toContain(
			'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>',
		);
		const reloaded = await loadDocx(saved);
		expect((reloaded.model.blocks[0] as Paragraph).runs.map((run) => run.text).join('')).toBe(
			'Intro1',
		);
	});
});
