import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index.js';
import { at, expectParagraph } from './test-support/access.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = `xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}"`;

// A minimal 1x1 transparent PNG, used only to exercise byte-preservation, never rendered.
const PNG_BASE64 =
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
const pngBytes = () => Uint8Array.from(Buffer.from(PNG_BASE64, 'base64'));

function inlinePicture(relId: string, descr = 'A test image'): string {
	return `<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1" descr="${descr}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

async function pictureFixture(documentXml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('word/document.xml', documentXml);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', pngBytes());
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('inline pictures', () => {
	it('parses an inline picture, its size/alt text, and exposes original bytes via media', async () => {
		const xml = `<w:document ${NS}><w:body><w:p><w:r>${inlinePicture('rId1')}</w:r><w:r><w:t> caption</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
		const loaded = await loadDocx(await pictureFixture(xml));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(at(paragraph.runs, 0).image).toMatchObject({
			relId: 'rId1',
			partName: 'word/media/image1.png',
			contentType: 'image/png',
			widthPx: 96,
			heightPx: 48,
			altText: 'A test image',
		});
		expect(at(paragraph.runs, 1).text).toBe(' caption');
		expect(loaded.media?.get('word/media/image1.png')).toEqual(pngBytes());
		expect(loaded.model.warnings.some((w) => w.includes('Inline pictures'))).toBe(true);
	});

	it('preserves an unchanged picture verbatim while editing sibling text', async () => {
		const xml = `<w:document ${NS}><w:body><w:p><w:r>${inlinePicture('rId1')}</w:r><w:r><w:t> caption</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
		const loaded = await loadDocx(await pictureFixture(xml));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		at(paragraph.runs, 1).text = ' new caption';
		const saved = await JSZip.loadAsync(await loaded.save());
		const savedXml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(savedXml).toContain('r:embed="rId1"');
		expect(savedXml).toContain('descr="A test image"');
		expect(savedXml).toContain('new caption');
		const rels = (await saved.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
		expect(rels.match(/<Relationship /g)).toHaveLength(1);
	});

	it('removes a deleted picture entirely on save', async () => {
		const xml = `<w:document ${NS}><w:body><w:p><w:r>${inlinePicture('rId1')}</w:r></w:p><w:sectPr/></w:body></w:document>`;
		const loaded = await loadDocx(await pictureFixture(xml));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs = [{ text: 'Replaced' }];
		const saved = await JSZip.loadAsync(await loaded.save());
		const savedXml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(savedXml).not.toContain('<w:drawing');
		expect(savedXml).toContain('Replaced');
	});

	it('inserts a brand-new picture, adding its media part, relationship and content type', async () => {
		const xml = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
		const zip = new JSZip();
		zip.file('word/document.xml', xml);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		paragraph.runs.push({
			text: '',
			image: {
				relId: 'pending',
				partName: 'word/media/inserted.png',
				contentType: 'image/png',
				widthPx: 32,
				heightPx: 32,
				altText: 'Inserted picture',
			},
		});
		const pendingMedia = new Map([
			['word/media/inserted.png', { bytes: pngBytes(), contentType: 'image/png' }],
		]);
		const saved = await JSZip.loadAsync(await loaded.save(loaded.model, pendingMedia));
		expect(await saved.file('word/media/inserted.png')?.async('uint8array')).toEqual(pngBytes());
		const savedXml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(savedXml).toContain('<w:drawing');
		const rels = (await saved.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
		expect(rels).toContain('media/inserted.png');
		expect(rels).toContain('relationships/image');
		const contentTypes = (await saved.file('[Content_Types].xml')?.async('string')) ?? '';
		expect(contentTypes).toContain('Extension="png"');
		const reopened = await loadDocx(await saved.generateAsync({ type: 'uint8array' }));
		const reopenedParagraph = expectParagraph(reopened.model.blocks[0]);
		expect(reopenedParagraph.runs.at(-1)?.image).toMatchObject({
			partName: 'word/media/inserted.png',
			widthPx: 32,
			heightPx: 32,
		});
	});

	it('models a floating picture as an inline placeholder and a chart as an unsupported placeholder', async () => {
		const anchored = `<w:drawing><wp:anchor><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
		const chart = `<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="urn:c" r:id="rId2"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;
		const xml = `<w:document ${NS}><w:body><w:p><w:r>${anchored}</w:r></w:p><w:p><w:r>${chart}</w:r></w:p><w:sectPr/></w:body></w:document>`;
		const loaded = await loadDocx(await pictureFixture(xml));
		const floating = expectParagraph(loaded.model.blocks[0]);
		const chartBlock = expectParagraph(loaded.model.blocks[1]);
		expect(at(floating.runs, 0).image).toMatchObject({ anchored: true, relId: 'rId1' });
		expect(at(chartBlock.runs, 0).image).toMatchObject({ unsupported: 'Chart' });
		expect(loaded.model.warnings.some((w) => w.includes('Floating'))).toBe(true);
		expect(loaded.model.warnings.some((w) => w.includes('Chart'))).toBe(true);
	});
});

describe('bookmarks', () => {
	it('exposes bookmark names read-only and preserves them (at the paragraph boundary) through an edit', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${W}"><w:body><w:p><w:bookmarkStart w:id="0" w:name="Intro"/><w:r><w:t>Hello</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.bookmarks).toEqual(['Intro']);
		expect(loaded.model.warnings.some((w) => w.includes('Bookmark'))).toBe(true);
		at(paragraph.runs, 0).text = 'Hello there';
		const saved = await JSZip.loadAsync(await loaded.save());
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('w:name="Intro"');
		expect(xml).toContain('<w:bookmarkStart');
		expect(xml).toContain('<w:bookmarkEnd');
		expect(xml).toContain('Hello there');
		const reopened = await loadDocx(await saved.generateAsync({ type: 'uint8array' }));
		const reopenedParagraph = expectParagraph(reopened.model.blocks[0]);
		expect(reopenedParagraph.bookmarks).toEqual(['Intro']);
	});
});
