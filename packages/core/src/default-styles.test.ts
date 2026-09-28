import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx } from './index.js';

describe('new document styles', () => {
	it("gives new documents Word's modern defaults and saves them in styles.xml", async () => {
		const model = createDocument();
		expect(model.paragraphStyles?.docDefaults).toMatchObject({
			spacingAfterTwips: 160,
			lineSpacingTwips: 259,
			lineSpacingRule: 'auto',
		});
		expect(model.characterStyles?.docDefaults).toMatchObject({
			fontFamily: 'Calibri',
			fontSize: 11,
		});
		expect(Object.keys(model.paragraphStyles!.styles)).toEqual(
			expect.arrayContaining(['Normal', 'Heading1', 'Heading2', 'Heading3', 'Title', 'TOC1']),
		);
		const zip = await JSZip.loadAsync(await saveDocx(model));
		expect(await zip.file('word/styles.xml')!.async('string')).toContain('<w:docDefaults>');
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"',
		);
		expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
			'Target="styles.xml"',
		);
		const reloaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(reloaded.model.paragraphStyles?.styles).toEqual(model.paragraphStyles?.styles);
	});

	it('still refuses changed style catalogs in the standalone writer', async () => {
		const model = createDocument();
		model.paragraphStyles!.styles.Normal.formatting.spacingAfterTwips = 0;
		await expect(saveDocx(model)).rejects.toThrow(/built-in default styles/);
	});
});
