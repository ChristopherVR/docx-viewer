import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx, createDocument, type Paragraph } from './index.js';
import { at } from './test-support/access.js';

const SVG = new TextEncoder().encode(
	'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>',
);
const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

describe('SVG pictures', () => {
	it('writes an svgBlip beside the PNG fallback and reads both back', async () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [
				{
					text: '',
					image: {
						relId: '',
						partName: 'word/media/logo.png',
						svgPartName: 'word/media/logo.svg',
						contentType: 'image/png',
						widthPx: 40,
						heightPx: 40,
					},
				},
			],
		};
		const saved = await saveDocx(
			model,
			new Map([
				['word/media/logo.png', { bytes: PNG, contentType: 'image/png' }],
				['word/media/logo.svg', { bytes: SVG, contentType: 'image/svg+xml' }],
			]),
		);
		const zip = await JSZip.loadAsync(saved);
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toMatch(
			/<a:blip [^>]*r:embed="(rId\d+)"><a:extLst><a:ext uri="\{96DAC541-7B7A-43D3-8B79-37D633B846F1\}"><asvg:svgBlip [^>]*r:embed="(rId\d+)"/,
		);
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'Extension="svg" ContentType="image/svg+xml"',
		);
		const reloaded = await loadDocx(saved);
		const image = at((at(reloaded.model.blocks, 0) as Paragraph).runs, 0).image;
		expect(image).toMatchObject({
			partName: 'word/media/logo.png',
			svgPartName: 'word/media/logo.svg',
		});
		expect(reloaded.media?.get('word/media/logo.svg')).toEqual(SVG);
	});
});
