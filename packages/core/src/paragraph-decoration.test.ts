import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type DocumentModel, type Paragraph } from './index.js';
import { expectParagraph, must } from './test-support/access.js';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>
<w:p><w:pPr><w:pBdr><w:bottom w:val="double" w:sz="12" w:space="4" w:color="FF0000"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Decorated</w:t></w:r></w:p>
<w:p><w:r><w:t>Plain</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const documentXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
const paragraphs = (model: DocumentModel) => model.blocks.map((b) => expectParagraph(b));

describe('paragraph shading and borders', () => {
	it('reads existing shading and borders into the model', async () => {
		const { model } = await loadDocx(await fixture());
		const [first] = paragraphs(model);
		expect(first!.shadingFill).toBe('#FFFF00');
		expect(first!.borders?.bottom).toMatchObject({
			style: 'double',
			sizeEighthPoints: 12,
			spacePoints: 4,
		});
	});

	it('writes new shading and borders on a plain paragraph', async () => {
		const loaded = await loadDocx(await fixture());
		const model = structuredClone(loaded.model);
		const plain = paragraphs(model)[1] as Paragraph;
		plain.shadingFill = '#c0e6f5';
		plain.borders = {
			top: { style: 'single', sizeEighthPoints: 8 as never, color: '#0070c0' },
			left: { style: 'single' },
			between: { style: 'dashed', sizeEighthPoints: 4 as never },
		};
		const xml = await documentXml(await loaded.save(model));
		expect(xml).toContain('<w:shd w:val="clear" w:color="auto" w:fill="C0E6F5"/>');
		expect(xml).toContain('w:top w:val="single" w:sz="8" w:space="1" w:color="0070C0"');
		expect(xml).toContain('w:between w:val="dashed"');
		// CT_PBdr order: top, left, bottom, right, between
		expect(xml.indexOf('<w:top')).toBeLessThan(xml.lastIndexOf('<w:left'));
		expect(xml.lastIndexOf('<w:left')).toBeLessThan(xml.indexOf('<w:between'));
		const reloaded = await loadDocx(await loaded.save(model));
		const round = paragraphs(reloaded.model)[1]!;
		expect(round.shadingFill).toBe('#C0E6F5');
		expect(round.borders?.top).toMatchObject({ style: 'single', color: '#0070C0' });
	});

	it('changes and removes shading and borders on a decorated paragraph', async () => {
		const loaded = await loadDocx(await fixture());
		const changed = structuredClone(loaded.model);
		const first = paragraphs(changed)[0] as Paragraph;
		first.shadingFill = '#00ff00';
		delete first.borders;
		const xml = await documentXml(await loaded.save(changed));
		expect(xml).toContain('w:fill="00FF00"');
		expect(xml).not.toContain('w:fill="FFFF00"');
		expect(xml).not.toContain('<w:pBdr>');
		expect(xml).toContain('<w:jc w:val="center"/>');
		delete first.shadingFill;
		const cleared = await documentXml(await loaded.save(changed));
		expect(cleared).not.toContain('<w:shd');
	});

	it('leaves unchanged paragraphs byte-identical when another paragraph changes', async () => {
		const original = await fixture();
		const loaded = await loadDocx(original);
		const model = structuredClone(loaded.model);
		(paragraphs(model)[1] as Paragraph).shadingFill = '#ffcc00';
		const xml = await documentXml(await loaded.save(model));
		const source = await documentXml(original);
		const decorated = /<w:p><w:pPr><w:pBdr>[\s\S]*?<\/w:p>/.exec(source)![0];
		expect(xml).toContain(decorated);
	});

	it('writes a border with style none without size or colour', async () => {
		const loaded = await loadDocx(await fixture());
		const model = structuredClone(loaded.model);
		(paragraphs(model)[1] as Paragraph).borders = { bottom: { style: 'none' } };
		const xml = await documentXml(await loaded.save(model));
		expect(xml).toContain('<w:bottom w:val="none"/>');
		expect(must(xml, 'xml')).not.toContain('w:sz="4" w:space="1" w:color="auto"/></w:pBdr>');
	});
});
