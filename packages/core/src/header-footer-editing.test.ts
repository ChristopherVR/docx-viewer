import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type DocumentModel, type Paragraph } from './index.js';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Two sections sharing header1.xml, plus a footer with a PAGE field. */
async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	const refs = `<w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:pPr><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${r}/footer" Target="footer1.xml"/></Relationships>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${w}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Draft</w:t></w:r></w:p></w:hdr>`,
	);
	zip.file(
		'word/footer1.xml',
		`<w:ftr xmlns:w="${w}"><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

/** Applies `edit` to every section's default header (they all reference one part). */
function editHeader(model: DocumentModel, edit: (paragraph: Paragraph) => void) {
	for (const section of model.sections ?? [])
		edit(section.headers!.default!.blocks[0] as Paragraph);
}

describe('header and footer editing', () => {
	it('rewrites only the edited header part and keeps its formatting', async () => {
		const original = await fixture();
		const loaded = await loadDocx(original);
		expect(loaded.model.sections?.[0].headers?.default?.partName).toBe('word/header1.xml');
		editHeader(loaded.model, (paragraph) => {
			paragraph.runs[0].text = 'Final';
		});
		const saved = await loaded.save(loaded.model);
		const zip = await JSZip.loadAsync(saved);
		const header = await zip.file('word/header1.xml')!.async('string');
		expect(header).toContain('Final');
		expect(header).toContain('<w:jc w:val="right"/>');
		expect(header).toContain('<w:b/>');
		const originalZip = await JSZip.loadAsync(original);
		expect(await zip.file('word/footer1.xml')!.async('string')).toBe(
			await originalZip.file('word/footer1.xml')!.async('string'),
		);
		const reloaded = await loadDocx(saved);
		const paragraph = reloaded.model.sections![1].headers!.default!.blocks[0] as Paragraph;
		expect(paragraph.runs[0]).toMatchObject({ text: 'Final', bold: true });
	});

	it('requires a shared header part to be edited consistently', async () => {
		const loaded = await loadDocx(await fixture());
		// Editors work on cloned models, where the sections' copies of a shared part can diverge.
		const model = JSON.parse(JSON.stringify(loaded.model)) as DocumentModel;
		(model.sections![0].headers!.default!.blocks[0] as Paragraph).runs[0].text = 'Only one';
		await expect(loaded.save(model)).rejects.toThrow(/shared by several sections/);
	});

	it('writes a new header link into the header part relationships', async () => {
		const loaded = await loadDocx(await fixture());
		editHeader(loaded.model, (paragraph) => {
			paragraph.runs[0] = { text: 'Site', link: { href: 'https://example.com/' } };
		});
		const zip = await JSZip.loadAsync(await loaded.save(loaded.model));
		const headerRels = await zip.file('word/_rels/header1.xml.rels')!.async('string');
		expect(headerRels).toContain('Target="https://example.com/"');
		expect(headerRels).toContain('TargetMode="External"');
		const header = await zip.file('word/header1.xml')!.async('string');
		const id = /<w:hyperlink[^>]*r:id="(rId\d+)"/.exec(header)?.[1];
		expect(headerRels).toContain(`Id="${id}"`);
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).not.toContain(
			'example.com',
		);
		const reloaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = reloaded.model.sections![0].headers!.default!.blocks[0] as Paragraph;
		expect(paragraph.runs[0]).toMatchObject({ link: { href: 'https://example.com/' } });
	});

	it('resolves pictures in a header through the header part relationships', async () => {
		const zip = await JSZip.loadAsync(await fixture());
		zip.file(
			'word/header1.xml',
			`<w:hdr xmlns:w="${w}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${r}"><w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Logo" descr="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>`,
		);
		zip.file(
			'word/_rels/header1.xml.rels',
			`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="${r}/image" Target="media/logo.png"/></Relationships>`,
		);
		zip.file('word/media/logo.png', new Uint8Array([137, 80, 78, 71]));
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.sections![0].headers!.default!.blocks[0] as Paragraph;
		expect(paragraph.runs[0].image).toMatchObject({
			partName: 'word/media/logo.png',
			widthPx: 100,
			altText: 'Logo',
		});
		expect(loaded.media?.get('word/media/logo.png')).toEqual(new Uint8Array([137, 80, 78, 71]));
	});

	it('writes page size and orientation for one section and protects other properties', async () => {
		const loaded = await loadDocx(await fixture());
		const model = JSON.parse(JSON.stringify(loaded.model)) as DocumentModel;
		Object.assign(model.sections![0], {
			pageWidthTwips: 15840,
			pageHeightTwips: 12240,
			orientation: 'landscape',
			columns: { count: 2, spacingTwips: 360, equalWidth: true },
		});
		const saved = await loaded.save(model);
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toMatch(
			/<w:pPr><w:sectPr>.*w:w="15840" w:h="12240" w:orient="landscape".*<w:cols w:num="2" w:space="360"\/>.*<\/w:sectPr><\/w:pPr>/,
		);
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.sections![0]).toMatchObject({
			orientation: 'landscape',
			columns: { count: 2 },
		});
		expect(reloaded.model.sections![1]).toMatchObject({ pageWidthTwips: 12240 });
		const titled = JSON.parse(JSON.stringify(loaded.model)) as DocumentModel;
		titled.sections![0].verticalAlign = 'center';
		await expect(loaded.save(titled)).rejects.toThrow(/verticalAlign/);
	});
});
