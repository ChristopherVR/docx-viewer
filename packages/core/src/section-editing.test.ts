import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type DocumentModel } from './index.js';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function docx(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('word/document.xml', `<w:document xmlns:w="${w}"><w:body>${body}</w:body></w:document>`);
	return zip.generateAsync({ type: 'uint8array' });
}
const documentXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
const clone = (model: DocumentModel) => JSON.parse(JSON.stringify(model)) as DocumentModel;
const finalSection =
	'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>';

describe('section editing', () => {
	it('inserts a continuous section break with two columns after a paragraph', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p>${finalSection}`,
			),
		);
		const model = clone(loaded.model);
		const [last] = model.sections!;
		model.sections = [
			{ ...structuredClone(last), endsAtBlockId: model.blocks[0].id },
			{ ...last, type: 'continuous', columns: { count: 2, spacingTwips: 720, equalWidth: true } },
		];
		const saved = await loaded.save(model);
		const xml = await documentXml(saved);
		expect(xml).toMatch(/Intro.*<\/w:p>.*Body.*<w:sectPr><w:type w:val="continuous"\/>/s);
		expect(xml).toMatch(/<w:p><w:pPr><w:sectPr>.*<\/w:sectPr><\/w:pPr><w:r><w:t>Intro/);
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.sections).toHaveLength(2);
		expect(reloaded.model.sections![1]).toMatchObject({
			type: 'continuous',
			columns: { count: 2 },
		});
	});

	it('removes a section break the model no longer has', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p>${finalSection}`,
			),
		);
		expect(loaded.model.sections).toHaveLength(2);
		const model = clone(loaded.model);
		model.sections = [model.sections![1]];
		const saved = await loaded.save(model);
		expect((await documentXml(saved)).match(/<w:sectPr>/g)).toHaveLength(1);
		expect((await loadDocx(saved)).model.sections).toHaveLength(1);
	});

	it('refuses to end a section on a table', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p/>${finalSection}`,
			),
		);
		const model = clone(loaded.model);
		const [last] = model.sections!;
		model.sections = [{ ...structuredClone(last), endsAtBlockId: model.blocks[0].id }, last];
		await expect(loaded.save(model)).rejects.toThrow(/must end on a paragraph/);
	});

	it('writes a different first page and restarted Roman page numbers', async () => {
		const loaded = await loadDocx(
			await docx(`<w:p><w:r><w:t>Text</w:t></w:r></w:p>${finalSection}`),
		);
		const model = clone(loaded.model);
		Object.assign(model.sections![0], {
			titlePage: true,
			pageNumbering: { start: 1, format: 'lowerRoman' },
		});
		const saved = await loaded.save(model);
		const xml = await documentXml(saved);
		expect(xml).toContain('<w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:cols');
		expect(xml).toMatch(/<w:cols[^>]*\/><w:titlePg\/>/);
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.sections![0]).toMatchObject({
			titlePage: true,
			pageNumbering: { start: 1, format: 'lowerRoman' },
		});
		const cleared = clone(reloaded.model);
		cleared.sections![0].titlePage = false;
		delete cleared.sections![0].pageNumbering;
		const again = await documentXml(await reloaded.save(cleared));
		expect(again).not.toContain('titlePg');
		expect(again).not.toContain('pgNumType');
	});
});
