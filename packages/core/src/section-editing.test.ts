import { twips } from './units.js';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type DocumentModel } from './index.js';
import { at } from './test-support/access.js';

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
	it('does not copy a following footer reference into an earlier new section without a footer', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>${finalSection}`,
			),
		);
		const next = clone(loaded.model);
		const section = next.sections![0]!;
		next.sections = [
			{ ...section, endsAtBlockId: 'p0' },
			{
				...section,
				endsAtBlockId: 'p1',
				footers: {
					default: {
						partName: 'word/footer1.xml',
						blocks: [
							{
								type: 'paragraph',
								id: 'f',
								runs: [{ text: '1', field: { instr: ' PAGE ', simple: true } }],
							},
						],
					},
				},
			},
		];
		const xml = await documentXml(await loaded.save(next));
		expect(xml.match(/<w:footerReference/g)).toHaveLength(1);
		expect(xml.indexOf('<w:footerReference')).toBeGreaterThan(xml.indexOf('Second'));
	});
	it('writes unequal widths, removes stale columns and preserves unmodeled column attributes', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:r><w:t>Columns</w:t></w:r></w:p>${finalSection.replace('<w:cols w:space="720"/>', '<w:cols w:equalWidth="0" w:num="3" w:space="720"><w:col w:w="2400" w:space="720" data-extra="keep"/><w:col w:w="2400" w:space="720"/><w:col w:w="3120"/></w:cols>')}`,
			),
		);
		const next = clone(loaded.model);
		next.sections![0]!.columns = {
			count: 2,
			equalWidth: false,
			widths: [{ widthTwips: twips(2880), spacingTwips: twips(720) }, { widthTwips: twips(5760) }],
		};
		const bytes = await loaded.save(next);
		const xml = await documentXml(bytes);
		expect(xml).toContain('w:equalWidth="0"');
		expect(xml).toContain('data-extra="keep"');
		expect(xml.match(/<w:col /g)).toHaveLength(2);
		expect((await loadDocx(bytes)).model.sections![0]!.columns).toMatchObject({
			count: 2,
			equalWidth: false,
			widths: [{ widthTwips: 2880, spacingTwips: 720 }, { widthTwips: 5760 }],
		});
		next.sections![0]!.columns = { count: 1, equalWidth: true };
		const cleared = await documentXml(await loaded.save(next));
		expect(cleared).not.toContain('<w:col ');
		expect(cleared).not.toContain('w:equalWidth');
	});
	it('inserts a continuous section break with two columns after a paragraph', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p>${finalSection}`,
			),
		);
		const model = clone(loaded.model);
		const last = at(model.sections, 0);
		model.sections = [
			{ ...structuredClone(last), endsAtBlockId: at(model.blocks, 0).id },
			{
				...last,
				type: 'continuous',
				columns: { count: 2, spacingTwips: twips(720), equalWidth: true },
			},
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
		model.sections = [at(model.sections, 1)];
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
		const last = at(model.sections, 0);
		model.sections = [{ ...structuredClone(last), endsAtBlockId: at(model.blocks, 0).id }, last];
		await expect(loaded.save(model)).rejects.toThrow(/must end on a paragraph/);
	});

	it('writes a different first page and restarted Roman page numbers', async () => {
		const loaded = await loadDocx(
			await docx(`<w:p><w:r><w:t>Text</w:t></w:r></w:p>${finalSection}`),
		);
		const model = clone(loaded.model);
		Object.assign(at(model.sections, 0), {
			titlePage: true,
			pageNumbering: { start: 1, format: 'lowerRoman' },
		});
		const saved = await loaded.save(model);
		const xml = await documentXml(saved);
		expect(xml).toContain('<w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:cols');
		expect(xml).toMatch(/<w:cols[^>]*\/><w:titlePg\/>/);
		const reloaded = await loadDocx(saved);
		expect(at(reloaded.model.sections, 0)).toMatchObject({
			titlePage: true,
			pageNumbering: { start: 1, format: 'lowerRoman' },
		});
		const cleared = clone(reloaded.model);
		const clearedSection = at(cleared.sections, 0);
		clearedSection.titlePage = false;
		delete clearedSection.pageNumbering;
		const again = await documentXml(await reloaded.save(cleared));
		expect(again).not.toContain('titlePg');
		expect(again).not.toContain('pgNumType');
	});

	it('writes vertical alignment and an odd-page section break', async () => {
		const loaded = await loadDocx(
			await docx(
				`<w:p><w:r><w:t>Title</w:t></w:r></w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p>${finalSection}`,
			),
		);
		const model = clone(loaded.model);
		const last = at(model.sections, 0);
		model.sections = [
			{ ...structuredClone(last), endsAtBlockId: at(model.blocks, 0).id, verticalAlign: 'center' },
			{ ...last, type: 'oddPage' },
		];
		const xml = await documentXml(await loaded.save(model));
		expect(xml).toMatch(/<w:pPr><w:sectPr>.*<w:vAlign w:val="center"\/>.*<\/w:sectPr><\/w:pPr>/);
		expect(xml).toContain('<w:type w:val="oddPage"/>');
		const reloaded = await loadDocx(await loaded.save(model));
		expect(
			reloaded.model.sections!.map((section) => [section.type, section.verticalAlign]),
		).toEqual([
			['nextPage', 'center'],
			['oddPage', undefined],
		]);
	});
});
