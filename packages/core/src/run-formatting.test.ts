import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	loadDocx,
	parseParagraphStyleCatalog,
	parseRunStyleCatalog,
	resolveRunFormatting,
} from './index.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const stylesXml = `<w:styles xmlns:w="${ns}">
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading"><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="112233"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:b/><w:i/></w:rPr></w:style>
<w:style w:type="character" w:styleId="EmphDerived"><w:basedOn w:val="Emph"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>`;

function catalog() {
	return parseRunStyleCatalog(stylesXml);
}

describe('character style catalog and run formatting resolution', () => {
	it('parses rPrDefault and character/paragraph style run properties', () => {
		const cat = catalog();
		expect(cat.docDefaults).toMatchObject({ fontSize: 11, fontFamily: 'Calibri' });
		expect(cat.styles.Emph).toMatchObject({
			type: 'character',
			formatting: { bold: true, italic: true },
		});
		expect(cat.styles.EmphDerived).toMatchObject({
			type: 'character',
			basedOn: 'Emph',
			formatting: { fontSize: 16 },
		});
		expect(cat.styles.Heading).toMatchObject({
			type: 'paragraph',
			formatting: { color: '#112233' },
		});
	});

	it('resolves docDefaults, paragraph style run props, character style chain and direct overrides', () => {
		const runCatalog = catalog();
		const paragraphCatalog = parseParagraphStyleCatalog(stylesXml);
		const resolved = resolveRunFormatting(
			{ text: 'x', style: 'EmphDerived' },
			{ paragraphStyleId: 'Heading', paragraphCatalog, runCatalog },
		);
		// docDefaults fontSize (11) is overridden by the character style chain's leaf (16).
		expect(resolved.fontSize).toBe(16);
		expect(resolved.fontFamily).toBe('Calibri');
		// Paragraph style ("Heading") run color carries through when no later level overrides it.
		expect(resolved.color).toBe('#112233');
		// bold/italic: only the character style chain (Emph) defines them => a single toggle => on.
		expect(resolved.bold).toBe(true);
		expect(resolved.italic).toBe(true);
	});

	it('applies Word toggle-property semantics (ECMA-376 §17.7.3)', () => {
		const styles = `<w:styles xmlns:w="${ns}">
<w:style w:type="paragraph" w:styleId="Heading"><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="character" w:styleId="A"><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="character" w:styleId="B"><w:basedOn w:val="A"/><w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="character" w:styleId="C"><w:basedOn w:val="A"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>
</w:styles>`;
		const runCatalog = parseRunStyleCatalog(styles);
		const paragraphCatalog = parseParagraphStyleCatalog(styles);
		const bold = (run: Parameters<typeof resolveRunFormatting>[0], paragraphStyleId?: string) =>
			resolveRunFormatting(run, {
				runCatalog,
				paragraphCatalog,
				...(paragraphStyleId !== undefined && { paragraphStyleId }),
			}).bold;
		// Within one style's basedOn chain values inherit: B and A both bold is still bold, and C
		// turning it off wins over A.
		expect(bold({ text: 'x', style: 'B' })).toBe(true);
		expect(bold({ text: 'x', style: 'C' })).toBeUndefined();
		// Across style types values XOR: a bold character style in a bold paragraph style is plain.
		expect(bold({ text: 'x', style: 'B' }, 'Heading')).toBeUndefined();
		// Direct formatting is absolute: on stays on, and an explicit off cancels the styles.
		expect(bold({ text: 'x', style: 'B', bold: true }, 'Heading')).toBe(true);
		expect(bold({ text: 'x', bold: false }, 'Heading')).toBeUndefined();
	});

	it('detects character style inheritance cycles without recursing forever', () => {
		const cat = parseRunStyleCatalog(
			`<w:styles xmlns:w="${ns}"><w:style w:type="character" w:styleId="A"><w:basedOn w:val="B"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="character" w:styleId="B"><w:basedOn w:val="A"/><w:rPr><w:i/></w:rPr></w:style></w:styles>`,
		);
		expect(cat.warnings[0]).toContain('inheritance cycle detected');
		expect(resolveRunFormatting({ text: 'x', style: 'A' }, { runCatalog: cat }).italic).toBe(true);
	});

	it('reads caps, smallCaps, doubleStrike, underline style/color, spacing, shading, vanish and rStyle end to end', async () => {
		const xml = `<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:rStyle w:val="Emph"/><w:caps/><w:smallCaps/><w:dstrike/><w:vanish/><w:u w:val="wave" w:color="FF0000"/><w:spacing w:val="20"/><w:shd w:val="clear" w:fill="ABCDEF"/></w:rPr><w:t>Styled</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
		const zip = new JSZip();
		zip.file('word/document.xml', xml);
		zip.file('word/styles.xml', stylesXml);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		const run = paragraph.runs[0];
		expect(run).toMatchObject({
			style: 'Emph',
			caps: true,
			smallCaps: true,
			doubleStrike: true,
			vanish: true,
			underline: true,
			underlineStyle: 'wave',
			underlineColor: '#FF0000',
			characterSpacingTwips: 20,
			shadingFill: '#ABCDEF',
		});
		const original = await zip.generateAsync({ type: 'uint8array' });
		const roundTripped = await loaded.save();
		expect(roundTripped).toEqual(original);
	});

	it('no longer reports a blanket rStyle warning and describes accurate limits instead', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:rPr><w:rStyle w:val="Emph"/></w:rPr><w:t>Styled</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file('word/styles.xml', stylesXml);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const warnings = loaded.model.warnings.join(' ');
		expect(warnings).not.toContain(
			'Character style inheritance and theme font/color resolution are not modeled',
		);
		expect(warnings).toContain('basedOn chains');
		// Editing the paragraph containing rStyle stays supported, and the style reference survives.
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		paragraph.runs[0].text = 'Edited styled';
		const saved = await JSZip.loadAsync(await loaded.save());
		const xml = await saved.file('word/document.xml')?.async('string');
		expect(xml).toContain('<w:rStyle w:val="Emph"');
		expect(xml).toContain('Edited styled');
	});
});
