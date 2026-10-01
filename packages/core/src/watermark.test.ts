import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type InlineImage, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WATERMARK_RUN =
	'<w:r><w:rPr><w:noProof/></w:rPr><w:pict><v:shapetype xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" id="_x0000_t136" coordsize="21600,21600" o:spt="136"/><v:shape xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" id="PowerPlusWaterMarkObject357" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;width:412pt;height:165pt;rotation:315;z-index:-251658752" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>';

async function fixture(headerBody: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/></Relationships>`,
	);
	zip.file(
		'word/header1.xml',
		`<w:hdr xmlns:w="${W}"><w:p>${headerBody}<w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const headerOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/header1.xml')?.async('string')) ?? '';
const watermarkRun = (loaded: Awaited<ReturnType<typeof loadDocx>>) =>
	(loaded.model.sections![0]!.headers!.default!.blocks[0] as Paragraph).runs.find(
		(run) => run.image?.watermark,
	);

describe('text watermarks', () => {
	it('reads a Word watermark as a labelled run with its settings', async () => {
		const loaded = await loadDocx(await fixture(WATERMARK_RUN));
		const image = watermarkRun(loaded)!.image!;
		expect(image).toMatchObject({
			unsupported: 'Watermark',
			watermark: {
				text: 'DRAFT',
				color: '#c0c0c0',
				semitransparent: true,
				layout: 'diagonal',
				fontFamily: 'Calibri',
			},
		});
	});

	it('keeps the original shape when only the header text changes', async () => {
		const bytes = await fixture(WATERMARK_RUN);
		const loaded = await loadDocx(bytes);
		const paragraph = loaded.model.sections![0]!.headers!.default!.blocks[0] as Paragraph;
		paragraph.runs.find((run) => run.text === 'Head')!.text = 'Header';
		const xml = await headerOf(await loaded.save());
		expect(xml).toContain('width:412pt;height:165pt;rotation:315');
		expect(xml).toContain('<w:t>Header</w:t>');
	});

	it('adds, replaces and removes a watermark run in the header paragraph', async () => {
		const loaded = await loadDocx(await fixture(''));
		const paragraph = loaded.model.sections![0]!.headers!.default!.blocks[0] as Paragraph;
		const image: InlineImage = {
			relId: '',
			partName: '',
			contentType: 'application/octet-stream',
			widthPx: 0,
			heightPx: 0,
			unsupported: 'Watermark',
			watermark: {
				text: 'CONFIDENTIAL',
				color: '#ff0000',
				semitransparent: false,
				layout: 'horizontal',
				fontFamily: 'Arial',
			},
		};
		paragraph.runs.unshift({ text: '', image });
		let saved = await loaded.save();
		let xml = await headerOf(saved);
		expect(xml).toMatch(/id="PowerPlusWaterMarkObject\d+"/);
		expect(xml).toContain('string="CONFIDENTIAL"');
		expect(xml).toContain('fillcolor="ff0000"');
		expect(xml).not.toContain('rotation:315');
		expect(xml).toContain('<w:t>Head</w:t>');
		const reopened = await loadDocx(saved);
		expect(watermarkRun(reopened)!.image!.watermark).toMatchObject({
			text: 'CONFIDENTIAL',
			color: '#ff0000',
			layout: 'horizontal',
			semitransparent: false,
		});
		watermarkRun(reopened)!.image!.watermark = {
			...watermarkRun(reopened)!.image!.watermark!,
			text: 'DRAFT',
			layout: 'diagonal',
		};
		xml = await headerOf(await reopened.save());
		expect(xml).toContain('string="DRAFT"');
		expect(xml).toContain('rotation:315');
		expect(xml.match(/PowerPlusWaterMarkObject/g)).toHaveLength(1);
		const again = await loadDocx(saved);
		const para = again.model.sections![0]!.headers!.default!.blocks[0] as Paragraph;
		para.runs = para.runs.filter((run) => !run.image);
		xml = await headerOf(await again.save());
		expect(xml).not.toContain('PowerPlusWaterMarkObject');
		expect(xml).toContain('<w:t>Head</w:t>');
	});
});
