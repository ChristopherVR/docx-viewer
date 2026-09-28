import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type DocumentModel, type Paragraph } from './index.js';
import { ensureListDefinition } from './numbering-editing.js';
import { at, lastOf } from './test-support/access.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

async function pack(files: Record<string, string | Uint8Array>): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [name, contents] of Object.entries(files)) zip.file(name, contents);
	return zip.generateAsync({ type: 'uint8array' });
}
const part = async (bytes: Uint8Array, name: string) =>
	(await JSZip.loadAsync(bytes)).file(name)!.async('string');
const clone = (model: DocumentModel) => structuredClone(model);
/** Local names of an element's direct children, from serialized XML that has no nested repeats. */
const childNames = (xml: string) => Array.from(xml.matchAll(/<w:(\w+)[ />]/g), (m) => m[1]);

describe('numbering.xml element order', () => {
	const document = `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

	async function withNewList(numberingXml: string): Promise<string> {
		const loaded = await loadDocx(
			await pack({ 'word/document.xml': document, 'word/numbering.xml': numberingXml }),
		);
		const model = clone(loaded.model);
		const { catalog, numId } = ensureListDefinition(model.numberingCatalog, 'decimal');
		const added = lastOf(Object.values(catalog.abstractNums));
		added.levels[0] = {
			...at(added.levels, 0),
			isLgl: true,
			lvlRestart: 0,
			lvlJc: 'right',
			suffix: 'space',
			indentLeftTwips: 720,
			hangingTwips: 360,
		};
		model.numberingCatalog = catalog;
		(model.blocks[0] as Paragraph).numbering = { numId, level: 0 };
		return part(await loaded.save(model), 'word/numbering.xml');
	}

	it('writes w:lvl children in CT_Lvl order', async () => {
		const xml = await withNewList(
			`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
		);
		const lvl = /<w:abstractNum w:abstractNumId="1">.*?<\/w:lvl>/s.exec(xml)![0];
		expect(childNames(lvl).filter((name) => name !== 'ind')).toEqual([
			'abstractNum',
			'multiLevelType',
			'lvl',
			'start',
			'numFmt',
			'lvlRestart',
			'isLgl',
			'suff',
			'lvlText',
			'lvlJc',
			'pPr',
		]);
	});

	it('inserts new definitions before numIdMacAtCleanup, abstractNums before nums', async () => {
		const xml = await withNewList(
			`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:numIdMacAtCleanup w:val="1"/></w:numbering>`,
		);
		const top = Array.from(
			xml.matchAll(/<w:(abstractNum|num|numIdMacAtCleanup)[ >/]/g),
			(m) => m[1],
		);
		expect(top).toEqual(['abstractNum', 'abstractNum', 'num', 'num', 'numIdMacAtCleanup']);
	});

	it('inserts a new abstractNum before numIdMacAtCleanup when the part has no w:num', async () => {
		const xml = await withNewList(
			`<w:numbering xmlns:w="${W}"><w:numIdMacAtCleanup w:val="0"/></w:numbering>`,
		);
		const top = Array.from(
			xml.matchAll(/<w:(abstractNum|num|numIdMacAtCleanup)[ >/]/g),
			(m) => m[1],
		);
		expect(top).toEqual(['abstractNum', 'num', 'numIdMacAtCleanup']);
	});
});

describe('picture extent updates', () => {
	const png = Uint8Array.from(
		Buffer.from(
			'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
			'base64',
		),
	);
	const drawing = `<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:blipFill><a:blip r:embed="rId1"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

	it('resizes wp:extent and a:xfrm/a:ext but leaves a:extLst/a:ext entries alone', async () => {
		const loaded = await loadDocx(
			await pack({
				'word/document.xml': `<w:document xmlns:w="${W}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}" xmlns:r="${R}"><w:body><w:p><w:r>${drawing}</w:r></w:p><w:sectPr/></w:body></w:document>`,
				'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/image" Target="media/image1.png"/></Relationships>`,
				'word/media/image1.png': png,
			}),
		);
		const model = clone(loaded.model);
		const image = at((at(model.blocks, 0) as Paragraph).runs, 0).image!;
		image.widthPx = 192;
		image.heightPx = 96;
		const xml = await part(await loaded.save(model), 'word/document.xml');
		expect(xml).toContain('<wp:extent cx="1828800" cy="914400"');
		expect(xml).toMatch(
			/<a:xfrm><a:off x="0" y="0"\/><a:ext cx="1828800" cy="914400"\/><\/a:xfrm>/,
		);
		const extLstEntry = /<a:ext uri="\{28A0092B[^>]*>/.exec(xml)![0];
		expect(extLstEntry).not.toMatch(/\bcx=|\bcy=/);
	});
});

describe('w:sectPr child order', () => {
	const refs = `<w:headerReference w:type="default" r:id="rId1"/>`;
	const relationships = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/></Relationships>`;
	const header = `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>`;

	it('inserts a missing pgSz and pgMar in sequence rather than first or last', async () => {
		const loaded = await loadDocx(
			await pack({
				'word/document.xml': `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr>${refs}<w:cols w:space="720"/><w:titlePg/></w:sectPr></w:body></w:document>`,
				'word/_rels/document.xml.rels': relationships,
				'word/header1.xml': header,
			}),
		);
		const model = clone(loaded.model);
		at((at(model.blocks, 0) as Paragraph).runs, 0).text = 'edited';
		const xml = await part(await loaded.save(model), 'word/document.xml');
		const sectPr = /<w:sectPr>.*<\/w:sectPr>/s.exec(xml)![0];
		expect(childNames(sectPr).filter((name) => name !== 'sectPr')).toEqual([
			'headerReference',
			'pgSz',
			'pgMar',
			'cols',
			'titlePg',
		]);
	});

	it('orders a new section break, and keeps sectPr ahead of pPrChange', async () => {
		const loaded = await loadDocx(
			await pack({
				'word/document.xml': `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:pPr><w:pPrChange w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr>${refs}<w:cols w:space="720"/></w:sectPr></w:body></w:document>`,
				'word/_rels/document.xml.rels': relationships,
				'word/header1.xml': header,
			}),
		);
		const model = clone(loaded.model);
		const last = at(model.sections, 0);
		model.sections = [
			{ ...structuredClone(last), endsAtBlockId: at(model.blocks, 0).id },
			{ ...last, type: 'continuous' },
		];
		const xml = await part(await loaded.save(model), 'word/document.xml');
		const paragraph = /<w:p><w:pPr>.*?<\/w:pPr>/s.exec(xml)![0];
		expect(paragraph.indexOf('<w:sectPr')).toBeGreaterThan(-1);
		expect(paragraph.indexOf('<w:sectPr')).toBeLessThan(paragraph.indexOf('<w:pPrChange'));
		const sectPr = /<w:sectPr>.*?<\/w:sectPr>/s.exec(paragraph)![0];
		expect(childNames(sectPr).filter((name) => name !== 'sectPr')).toEqual([
			'headerReference',
			'pgSz',
			'pgMar',
			'cols',
		]);
	});
});
