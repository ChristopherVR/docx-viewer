import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx, type Paragraph } from './index.js';
import { paragraphJustification } from './paragraph-alignment.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function packageFor(jc: string, extra = ''): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:pPr>${extra}<w:jc w:val="${jc}"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
async function documentXml(bytes: Uint8Array): Promise<string> {
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}
const paragraphOf = (model: { blocks: unknown[] }) => model.blocks[0] as Paragraph;

describe('justification on write', () => {
	it('round-trips start, end and distribute losslessly when the text is edited', async () => {
		for (const jc of ['start', 'end', 'distribute', 'thaiDistribute', 'numTab']) {
			const loaded = await loadDocx(await packageFor(jc));
			paragraphOf(loaded.model).runs[0].text = 'Edited';
			const xml = await documentXml(await loaded.save());
			expect(xml, jc).toContain(`<w:jc w:val="${jc}"`);
		}
	});

	it('writes the exact justification for a new paragraph whose align still matches it', async () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p',
			align: 'justify',
			justification: 'distribute',
			runs: [{ text: 'x' }],
		};
		expect(await documentXml(await saveDocx(model))).toContain('<w:jc w:val="distribute"');
	});

	it('writes align when an editor changed it and left justification stale', async () => {
		const loaded = await loadDocx(await packageFor('distribute'));
		paragraphOf(loaded.model).align = 'center';
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w:jc w:val="center"');
		expect(xml).not.toContain('distribute');
	});

	it('removes jc when align is cleared, even though justification is stale', async () => {
		const loaded = await loadDocx(await packageFor('end'));
		paragraphOf(loaded.model).align = undefined;
		expect(await documentXml(await loaded.save())).not.toContain('<w:jc');
	});

	it('resolves start/end against the paragraph direction', () => {
		const base: Paragraph = { type: 'paragraph', id: 'p', runs: [] };
		expect(
			paragraphJustification({ ...base, justification: 'start', align: 'right', direction: 'rtl' }),
		).toBe('start');
		expect(
			paragraphJustification({ ...base, justification: 'start', align: 'right', direction: 'ltr' }),
		).toBe('right');
		expect(paragraphJustification({ ...base, align: 'justify' })).toBe('both');
		expect(paragraphJustification({ ...base })).toBeUndefined();
	});
});
