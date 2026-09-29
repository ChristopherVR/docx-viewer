import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	createDocument,
	loadDocx,
	saveDocx,
	twips,
	type DocumentModel,
	type Paragraph,
} from './index.js';
import { at, expectParagraph, must } from './test-support/access.js';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A package with body text and a section, but no header or footer parts. */
async function plainPackage(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}" xmlns:r="${r}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${r}/styles" Target="styles.xml"/></Relationships>`,
	);
	zip.file('word/styles.xml', `<w:styles xmlns:w="${w}"/>`);
	return zip.generateAsync({ type: 'uint8array' });
}

const pageNumber = (id: string): Paragraph => ({
	type: 'paragraph',
	id,
	align: 'center',
	runs: [{ text: '1', field: { instr: ' PAGE ', simple: true } }],
});

describe('creating headers and footers', () => {
	it('adds a footer part, its relationship, content type and sectPr reference to a package', async () => {
		const loaded = await loadDocx(await plainPackage());
		const model = structuredClone(loaded.model);
		const section = must(model.sections?.[0], 'section');
		section.footers = { default: { partName: 'word/footer1.xml', blocks: [pageNumber('f1')] } };
		const saved = await loaded.save(model);
		const zip = await JSZip.loadAsync(saved);
		const footer = await zip.file('word/footer1.xml')!.async('string');
		expect(footer).toContain('<w:ftr');
		expect(footer).toContain('PAGE');
		expect(footer).toContain('w:jc w:val="center"');
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"',
		);
		const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		expect(rels).toMatch(/Id="(rId\d+)"[^>]*relationships\/footer"[^>]*Target="footer1.xml"/);
		const rid = /Id="(rId\d+)"[^>]*relationships\/footer"/.exec(rels)![1];
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toContain(`<w:footerReference w:type="default" r:id="${rid}"/>`);
		expect(rid).not.toBe('rId1');
		expect(await zip.file('word/styles.xml')!.async('string')).toBe(`<w:styles xmlns:w="${w}"/>`);

		const reloaded = await loadDocx(saved);
		const paragraph = expectParagraph(
			must(at(reloaded.model.sections, 0).footers?.default, 'footer').blocks[0],
		);
		expect(paragraph.align).toBe('center');
		expect(paragraph.runs[0]?.field?.instr.trim()).toBe('PAGE');
	});

	it('creates a header and a footer together with distinct relationship ids', async () => {
		const loaded = await loadDocx(await plainPackage());
		const model = structuredClone(loaded.model);
		const section = must(model.sections?.[0], 'section');
		section.headers = {
			default: { partName: 'word/header1.xml', blocks: [{ ...pageNumber('h1'), align: 'right' }] },
		};
		section.footers = { default: { partName: 'word/footer1.xml', blocks: [pageNumber('f1')] } };
		const zip = await JSZip.loadAsync(await loaded.save(model));
		const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((match) => match[1]);
		expect(new Set(ids).size).toBe(ids.length);
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toContain('<w:headerReference w:type="default"');
		expect(document).toContain('<w:footerReference w:type="default"');
	});

	it('leaves the package untouched when nothing was added', async () => {
		const original = await plainPackage();
		const loaded = await loadDocx(original);
		const saved = await loaded.save(structuredClone(loaded.model));
		expect(Buffer.from(saved).equals(Buffer.from(original))).toBe(true);
	});

	it('rejects a part name that could escape the word folder', async () => {
		const loaded = await loadDocx(await plainPackage());
		const model = structuredClone(loaded.model);
		must(model.sections?.[0], 'section').footers = {
			default: { partName: '../evil.xml', blocks: [pageNumber('f1')] },
		};
		await expect(loaded.save(model)).rejects.toThrow(/Unsupported footer part name/);
	});

	it('saves a new document with a footer instead of refusing it', async () => {
		const model: DocumentModel = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'New' }] }];
		model.sections = [
			{
				endsAtBlockId: 'p1',
				type: 'nextPage',
				pageWidthTwips: twips(12240),
				pageHeightTwips: twips(15840),
				orientation: 'portrait',
				marginTopTwips: twips(1440),
				marginRightTwips: twips(1440),
				marginBottomTwips: twips(1440),
				marginLeftTwips: twips(1440),
				columns: { count: 1, equalWidth: true },
				footers: { default: { partName: 'word/footer1.xml', blocks: [pageNumber('f1')] } },
			},
		];
		const saved = await saveDocx(model);
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file('word/footer1.xml')!.async('string')).toContain('PAGE');
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/word/footer1.xml');
		const reloaded = await loadDocx(saved);
		expect(at(reloaded.model.sections, 0).footers?.default?.partName).toBe('word/footer1.xml');
	});
});
