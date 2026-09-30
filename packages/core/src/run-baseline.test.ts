import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, saveDocx, resolveRunFormatting } from './index.js';
import { parseRunProperties } from './run-properties.js';
import { parseXml, WORD_NS } from './xml.js';
import { runHasUnknownProperties } from './write-run-validation.js';
import { expectParagraph } from './test-support/access.js';

describe('explicit run baseline', () => {
	it('parses baseline as an editable neutral override and protects invalid values', () => {
		const run = parseXml(
			`<w:r xmlns:w="${WORD_NS}"><w:rPr><w:vertAlign w:val="baseline"/></w:rPr><w:t>Normal</w:t></w:r>`,
		).documentElement;
		expect(runHasUnknownProperties(run)).toBe(false);
		expect(parseRunProperties(run.firstChild as Element).verticalAlign).toBe('baseline');
		const invalid = parseXml(
			`<w:r xmlns:w="${WORD_NS}"><w:rPr><w:vertAlign w:val="unknown"/></w:rPr></w:r>`,
		).documentElement;
		expect(runHasUnknownProperties(invalid)).toBe(true);
	});
	it.each(['superscript', 'subscript'] as const)(
		'resets inherited %s without flattening styles and can restore inheritance',
		async (script) => {
			const zip = new JSZip();
			zip.file(
				'word/styles.xml',
				`<w:styles xmlns:w="${WORD_NS}"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:vertAlign w:val="${script}"/><w:sz w:val="24"/></w:rPr></w:style></w:styles>`,
			);
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Original</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
			);
			const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
			const run = expectParagraph(model.blocks[0]).runs[0]!;
			const context = {
				runCatalog: model.characterStyles,
				paragraphCatalog: model.paragraphStyles,
			};
			expect(resolveRunFormatting(run, context).verticalAlign).toBe(script);
			run.verticalAlign = 'baseline';
			run.text = 'Reset';
			const output = await saveDocx(model);
			const reopened = await loadDocx(output);
			expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toMatchObject({
				text: 'Reset',
				bold: true,
				verticalAlign: 'baseline',
			});
			const saved = await JSZip.loadAsync(output);
			expect(await saved.file('word/styles.xml')!.async('string')).toBe(
				await zip.file('word/styles.xml')!.async('string'),
			);
			delete run.verticalAlign;
			expect(resolveRunFormatting(run, context).verticalAlign).toBe(script);
			const inherited = await JSZip.loadAsync(await saveDocx(model));
			expect(await inherited.file('word/document.xml')!.async('string')).not.toContain(
				'w:vertAlign',
			);
		},
	);
});
