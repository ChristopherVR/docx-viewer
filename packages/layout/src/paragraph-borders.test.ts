import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import { layoutDocument } from './layout.js';
import type { LayoutParagraph } from './input.js';
import type { LayoutParagraphBox } from './result.js';
import type { TextMeasurer } from './measure.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer: TextMeasurer = { widthOf: (text) => text.length * 10, lineHeightOf: () => 20 };
const box = `<w:pBdr><w:top w:val="single" w:sz="12" w:space="3" w:color="C00000"/><w:bottom w:val="single" w:sz="12" w:space="3" w:color="C00000"/><w:between w:val="dotted" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:fill="FFF2CC"/>`;

async function load() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>` +
			`<w:p><w:pPr>${box}</w:pPr><w:r><w:t>One</w:t></w:r></w:p>` +
			`<w:p><w:pPr>${box}</w:pPr><w:r><w:t>Two</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

describe('paragraph borders and shading', () => {
	it('parses w:pBdr with spacing and w:shd', async () => {
		const { model } = await load();
		expect((model.blocks[0] as Paragraph).borders?.top).toMatchObject({
			style: 'single',
			sizeEighthPoints: 12,
			spacePoints: 3,
			color: '#C00000',
		});
		expect((model.blocks[0] as Paragraph).shadingFill).toBe('#FFF2CC');
	});

	it('draws a group of identical bordered paragraphs as one box with the between line', async () => {
		const { model } = await load();
		const [one, two, after] = adaptDocumentModel(model).sections[0].blocks as LayoutParagraph[];
		expect(one.borders?.top?.color).toBe('#C00000');
		expect(one.borders?.bottom).toBeUndefined();
		expect(two.borders?.top).toMatchObject({ style: 'dotted' });
		expect(two.borders?.bottom?.color).toBe('#C00000');
		expect(after.borders).toBeUndefined();
		const result = layoutDocument(adaptDocumentModel(model), measurer);
		const [first] = result.pages[0].columns[0].blocks as LayoutParagraphBox[];
		// 12 eighths of a point = 2px line, plus 3pt = 4px of space above the text.
		expect(first.lines[0].yPx).toBe(6);
		expect(first.frame).toMatchObject({ shading: '#FFF2CC' });
	});
});
