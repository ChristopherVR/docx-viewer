import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function load(pPr: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>O</w:t></w:r></w:p><w:p><w:r><w:t>nce upon a time</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const documentXml = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')) as string;

describe('drop cap frames', () => {
	it('reads w:framePr dropCap and lines', async () => {
		const loaded = await load('<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around"/>');
		expect((loaded.model.blocks[0] as Paragraph).dropCap).toEqual({ style: 'drop', lines: 3 });
	});

	it('writes a new drop cap and keeps a loaded frame untouched', async () => {
		const plain = await load('');
		const first = plain.model.blocks[0] as Paragraph;
		const xml = await documentXml(
			await plain.save({
				...plain.model,
				blocks: [
					{ ...first, dropCap: { style: 'margin', lines: 2 } },
					...plain.model.blocks.slice(1),
				],
			}),
		);
		expect(xml).toMatch(
			/<w:framePr [^>]*w:dropCap="margin"[^>]*w:lines="2"|<w:framePr [^>]*w:lines="2"[^>]*w:dropCap="margin"/,
		);

		const framed = await load(
			'<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:hSpace="20"/>',
		);
		const kept = await documentXml(await framed.save(framed.model));
		expect(kept).toContain('w:hSpace="20"');
		const removed = await documentXml(
			await framed.save({
				...framed.model,
				blocks: [
					(({ dropCap: _cap, ...rest }) => rest)(framed.model.blocks[0] as Paragraph),
					...framed.model.blocks.slice(1),
				],
			}),
		);
		expect(removed).not.toContain('framePr');
	});

	it('rejects an invalid line count before saving', async () => {
		const plain = await load('');
		const first = plain.model.blocks[0] as Paragraph;
		await expect(
			plain.save({ ...plain.model, blocks: [{ ...first, dropCap: { style: 'drop', lines: 0 } }] }),
		).rejects.toThrow(/dropCap/);
	});
});
