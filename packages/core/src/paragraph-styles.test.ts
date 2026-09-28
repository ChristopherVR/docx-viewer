import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, parseParagraphStyleCatalog, resolveParagraphFormatting } from './index.js';
import { at, expectParagraph, must } from './test-support/access.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const stylesXml = `<w:styles xmlns:w="${ns}">
<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="120"/><w:jc w:val="left"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="240"/><w:ind w:left="120"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Derived"><w:name w:val="Derived"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:line="360" w:lineRule="exact"/><w:bidi/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Char"><w:rPr><w:i/></w:rPr></w:style>
</w:styles>`;
const docXml = `<w:document xmlns:w="${ns}"><w:body>
<w:p><w:pPr><w:pStyle w:val="Derived"/><w:spacing w:after="480"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p>
<w:p><w:r><w:t>Default style</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:before="60"/></w:pPr><w:r><w:t>Direct</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>`;

async function fixture(styles = stylesXml, documentXml = docXml) {
	const zip = new JSZip();
	zip.file('word/document.xml', documentXml);
	zip.file('word/styles.xml', styles);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('paragraph style catalog', () => {
	it('resolves defaults, basedOn chains and direct overrides independently', async () => {
		const loaded = await loadDocx(await fixture());
		const [styled, defaulted, direct] = loaded.model.blocks;
		if (
			styled?.type !== 'paragraph' ||
			defaulted?.type !== 'paragraph' ||
			direct?.type !== 'paragraph'
		)
			throw new Error('Expected paragraph fixtures');
		expect(loaded.model.paragraphStyles?.styles.Derived).toMatchObject({
			name: 'Derived',
			basedOn: 'Normal',
			formatting: { direction: 'rtl', lineSpacingTwips: 360, lineSpacingRule: 'exact' },
		});
		expect(resolveParagraphFormatting(styled, loaded.model.paragraphStyles!)).toMatchObject({
			align: 'left',
			direction: 'rtl',
			spacingAfterTwips: 480,
			indentLeftTwips: 120,
			lineSpacingTwips: 360,
			lineSpacingRule: 'exact',
		});
		expect(resolveParagraphFormatting(defaulted, loaded.model.paragraphStyles!)).toMatchObject({
			align: 'left',
			spacingAfterTwips: 240,
			indentLeftTwips: 120,
		});
		expect(resolveParagraphFormatting(direct, loaded.model.paragraphStyles!)).toMatchObject({
			align: 'left',
			spacingBeforeTwips: 60,
			spacingAfterTwips: 240,
		});
		expect(styled).toMatchObject({ style: 'Derived', spacingAfterTwips: 480 });
		expect(styled.align).toBeUndefined();
		expect(loaded.model.warnings.join(' ')).toContain(
			'Paragraph style, character style and docDefaults inheritance resolve for rendering',
		);
	});

	it('returns original bytes unchanged and does not flatten inherited values after a text edit', async () => {
		const bytes = await fixture();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		at(paragraph.runs, 0).text = 'Edited styled paragraph';
		const saved = new Uint8Array(await loaded.save());
		const output = await JSZip.loadAsync(saved);
		expect(await output.file('word/styles.xml')?.async('string')).toBe(stylesXml);
		const xml = await output.file('word/document.xml')?.async('string');
		expect(xml).toContain('<w:pStyle w:val="Derived"');
		expect(xml).toContain('<w:spacing w:after="480"');
		expect(xml).not.toContain('w:line="360"');
	});

	it('reports inheritance cycles and resolves them without recursion failure', async () => {
		const cyclic = `<w:styles xmlns:w="${ns}"><w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style></w:styles>`;
		const loaded = await loadDocx(
			await fixture(
				cyclic,
				`<w:document xmlns:w="${ns}"><w:body><w:p><w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>Cycle</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
			),
		);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(loaded.model.paragraphStyles?.warnings[0]).toContain('inheritance cycle detected');
		expect(resolveParagraphFormatting(paragraph, loaded.model.paragraphStyles!).align).toBe(
			'center',
		);
	});

	it('rejects catalog mutations because styles.xml editing is unsupported', async () => {
		const loaded = await loadDocx(await fixture());
		must(loaded.model.paragraphStyles?.styles.Derived, 'Derived style').formatting.align = 'center';
		await expect(loaded.save()).rejects.toThrow(
			'Editing the paragraph style catalog is not supported',
		);
	});

	it('handles prototype-like style ids and long inheritance chains safely', () => {
		const catalog = parseParagraphStyleCatalog(
			`<w:styles xmlns:w="${ns}"><w:style w:type="paragraph" w:styleId="__proto__"><w:pPr><w:jc w:val="right"/></w:pPr></w:style></w:styles>`,
		);
		expect(Object.hasOwn(catalog.styles, '__proto__')).toBe(true);
		const styles = Object.create(null) as typeof catalog.styles;
		for (let index = 0; index < 4000; index++) {
			const id = `s${index}`;
			styles[id] = {
				id,
				...(index && { basedOn: `s${index - 1}` }),
				formatting: index === 0 ? { align: 'center' } : {},
			};
		}
		const deep = { ...catalog, styles };
		const paragraph = { type: 'paragraph' as const, id: 'p', style: 's3999', runs: [{ text: '' }] };
		expect(resolveParagraphFormatting(paragraph, deep).align).toBe('center');
	});
});
