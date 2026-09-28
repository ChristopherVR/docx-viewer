// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { DOMParser } from 'prosemirror-model';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { schema } from './schema';

function markAttrs(html: string, name: string) {
	const element = document.createElement('div');
	element.innerHTML = html;
	const doc = DOMParser.fromSchema(schema).parse(element);
	return doc.firstChild?.firstChild?.marks.find((mark) => mark.type.name === name)?.attrs;
}

describe('pasted styles', () => {
	it('keeps one font name and a hex color, as Word stores them', () => {
		expect(
			markAttrs(
				`<p><span style="font-family: 'Segoe UI', Arial, sans-serif; color: rgb(255, 0, 16); font-size: 16px">x</span></p>`,
				'font',
			),
		).toEqual({ family: 'Segoe UI', size: 12, color: '#ff0010' });
		expect(markAttrs('<p><span style="color: #abc">x</span></p>', 'font')).toEqual({
			family: null,
			size: null,
			color: '#aabbcc',
		});
		// Generic families and transparent colors carry no Word formatting.
		expect(
			markAttrs(
				'<p><span style="font-family: sans-serif; color: rgba(0, 0, 0, 0)">x</span></p>',
				'font',
			),
		).toBeUndefined();
		expect(
			markAttrs('<p><span style="background-color: rgb(255, 255, 255)">x</span></p>', 'highlight'),
		).toBeUndefined();
	});

	it('never writes a color that is not RRGGBB', async () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{ text: 'a', color: 'rgb(1, 2, 3)' },
					{ text: 'b', color: '#A1B2C3' },
				],
			},
		];
		// Pre-save validation rejects the model instead of silently dropping the bad color.
		await expect(saveDocx(model)).rejects.toThrow(/runs\[0\]\.color.*ST_HexColor/);
		const [paragraph] = model.blocks;
		if (paragraph.type === 'paragraph') paragraph.runs.shift();
		const xml = await (
			await JSZip.loadAsync(await saveDocx(model))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).not.toContain('rgb(');
		expect(xml).toContain('<w:color w:val="A1B2C3"/>');
	});
});
