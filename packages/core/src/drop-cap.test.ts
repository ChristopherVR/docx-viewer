import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, twips, type Paragraph } from './index.js';

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
	it('reads and edits distance without losing unrelated imported frame attributes', async () => {
		const loaded = await load(
			'<w:framePr w:dropCap="drop" w:lines="4" w:hSpace="144" w:vAnchor="text" w:x="80"/>',
		);
		const first = loaded.model.blocks[0] as Paragraph;
		expect(first.dropCap).toEqual({ style: 'drop', lines: 4, distanceTwips: 144 });
		const xml = await documentXml(
			await loaded.save({
				...loaded.model,
				blocks: [
					{ ...first, dropCap: { ...first.dropCap!, distanceTwips: twips(288) } },
					...loaded.model.blocks.slice(1),
				],
			}),
		);
		expect(xml).toContain('w:hSpace="288"');
		expect(xml).toContain('w:x="80"');
		expect(xml).toContain('w:vAnchor="text"');
		const cleared = await documentXml(
			await loaded.save({
				...loaded.model,
				blocks: [
					{ ...first, dropCap: { style: 'drop', lines: 4 } },
					...loaded.model.blocks.slice(1),
				],
			}),
		);
		expect(cleared).not.toContain('w:hSpace=');
	});
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
