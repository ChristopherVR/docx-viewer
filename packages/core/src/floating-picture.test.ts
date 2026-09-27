import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS = `xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

const anchor = (wrap: string, position: string, extra = 'behindDoc="0"') =>
	`<w:r><w:drawing><wp:anchor ${extra}><wp:positionH relativeFrom="column">${position}</wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="952500" cy="952500"/>${wrap}<wp:docPr id="1" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

async function fixture(body: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
	);
	zip.file('word/media/image1.png', new Uint8Array([137, 80, 78, 71]));
	return zip.generateAsync({ type: 'uint8array' });
}

describe('floating pictures', () => {
	it('parses wrapping and horizontal position from wp:anchor', async () => {
		const loaded = await loadDocx(
			await fixture(
				`<w:p>${anchor('<wp:wrapSquare wrapText="bothSides"/>', '<wp:align>right</wp:align>')}<w:r><w:t>Text</w:t></w:r></w:p>` +
					`<w:p>${anchor('<wp:wrapTopAndBottom/>', '<wp:posOffset>1905000</wp:posOffset>')}</w:p>` +
					`<w:p>${anchor('<wp:wrapNone/>', '<wp:align>left</wp:align>', 'behindDoc="1"')}</w:p>`,
			),
		);
		const images = loaded.model.blocks.map((block) => (block as Paragraph).runs[0].image);
		expect(images[0]).toMatchObject({
			anchored: true,
			placement: { wrap: 'square', align: 'right', relativeFrom: 'column' },
		});
		expect(images[1]?.placement).toMatchObject({ wrap: 'topAndBottom', offsetXPx: 200 });
		expect(images[2]?.placement).toMatchObject({ wrap: 'none', behindText: true });
		expect(images[0]?.placement).toMatchObject({ relativeFromV: 'paragraph', offsetYPx: 0 });
	});

	it('parses vertical alignment and offsets from wp:positionV', async () => {
		const xml = anchor('<wp:wrapNone/>', '<wp:posOffset>0</wp:posOffset>').replace(
			'<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>',
			'<wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>',
		);
		const loaded = await loadDocx(await fixture(`<w:p>${xml}</w:p>`));
		const image = (loaded.model.blocks[0] as Paragraph).runs[0].image;
		expect(image?.placement).toMatchObject({ relativeFromV: 'page', alignV: 'bottom' });
		const offset = anchor('<wp:wrapNone/>', '<wp:posOffset>0</wp:posOffset>').replace(
			'<wp:posOffset>0</wp:posOffset></wp:positionV>',
			'<wp:posOffset>476250</wp:posOffset></wp:positionV>',
		);
		const shifted = await loadDocx(await fixture(`<w:p>${offset}</w:p>`));
		expect((shifted.model.blocks[0] as Paragraph).runs[0].image?.placement?.offsetYPx).toBe(50);
	});

	it('keeps the anchor XML when text beside a floating picture is edited', async () => {
		const loaded = await loadDocx(
			await fixture(
				`<w:p>${anchor('<wp:wrapSquare wrapText="bothSides"/>', '<wp:align>right</wp:align>')}<w:r><w:t>Text</w:t></w:r></w:p>`,
			),
		);
		(loaded.model.blocks[0] as Paragraph).runs[1].text = 'Edited text';
		const xml = await (
			await JSZip.loadAsync(await loaded.save(loaded.model))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toContain('<wp:wrapSquare wrapText="bothSides"/>');
		expect(xml).toContain('<wp:align>right</wp:align>');
		expect(xml).toContain('Edited text');
	});
});
