import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, resolveParagraphFormatting, type Paragraph } from './index.js';
import { at, expectTable } from './test-support/access.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function load(body: string, extra: Record<string, string> = {}) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`,
	);
	for (const [path, xml] of Object.entries(extra)) zip.file(path, xml);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}
const paragraph = (loaded: Awaited<ReturnType<typeof load>>, index = 0): Paragraph => {
	const block = loaded.model.blocks[index];
	if (block?.type !== 'paragraph') throw new Error('Expected paragraph');
	return block;
};
const p = (pPr: string) => `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>x</w:t></w:r></w:p>`;

describe('paragraph justification', () => {
	it('keeps the exact ST_Jc value and derives the renderer alignment', async () => {
		const loaded = await load(
			[
				'<w:jc w:val="start"/>',
				'<w:jc w:val="end"/>',
				'<w:jc w:val="distribute"/>',
				'<w:jc w:val="thaiDistribute"/>',
				'<w:jc w:val="both"/>',
				'<w:jc w:val="numTab"/>',
				'<w:bidi/><w:jc w:val="start"/>',
				'<w:bidi/><w:jc w:val="end"/>',
			]
				.map(p)
				.join('') + '<w:sectPr/>',
		);
		const seen = loaded.model.blocks.map((_, index) => {
			const { align, justification } = paragraph(loaded, index);
			return [justification, align];
		});
		expect(seen).toEqual([
			['start', 'left'],
			['end', 'right'],
			['distribute', 'justify'],
			['thaiDistribute', 'justify'],
			['both', 'justify'],
			['numTab', undefined],
			['start', 'right'],
			['end', 'left'],
		]);
	});

	it('uses one parser for styles and paragraphs, dropping invalid values with a warning', async () => {
		const loaded = await load(
			p('<w:pStyle w:val="S"/>') + p('<w:jc w:val="justify"/>') + '<w:sectPr/>',
			{
				'word/styles.xml': `<w:styles xmlns:w="${ns}"><w:style w:type="paragraph" w:styleId="S"><w:pPr><w:jc w:val="end"/></w:pPr></w:style></w:styles>`,
			},
		);
		expect(loaded.model.paragraphStyles?.styles.S?.formatting).toMatchObject({
			align: 'right',
			justification: 'end',
		});
		expect(paragraph(loaded, 1).align).toBeUndefined();
		expect(paragraph(loaded, 1).justification).toBeUndefined();
		expect(loaded.model.warnings.join('\n')).toContain('Ignored invalid w:jc value “justify”');
	});

	it('resolves start/end against a direction inherited from the style', async () => {
		const loaded = await load(p('<w:pStyle w:val="S"/><w:jc w:val="start"/>') + '<w:sectPr/>', {
			'word/styles.xml': `<w:styles xmlns:w="${ns}"><w:style w:type="paragraph" w:styleId="S"><w:pPr><w:bidi/></w:pPr></w:style></w:styles>`,
		});
		const resolved = resolveParagraphFormatting(paragraph(loaded), loaded.model.paragraphStyles!);
		expect(resolved).toMatchObject({ direction: 'rtl', justification: 'start', align: 'right' });
	});

	it('round-trips an unedited distribute value unchanged', async () => {
		const loaded = await load(p('<w:jc w:val="distribute"/>') + '<w:sectPr/>');
		const saved = await JSZip.loadAsync(await loaded.save());
		expect(await saved.file('word/document.xml')?.async('string')).toContain('w:val="distribute"');
	});
});

describe('table jc and cell vertical alignment', () => {
	const table = (tblPr: string, tcPr = '') =>
		`<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tr><w:tc><w:tcPr>${tcPr}</w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
	it('accepts start/end (ST_JcTable), honouring bidiVisual, and vAlign both', async () => {
		const loaded = await load(
			table('<w:jc w:val="end"/>', '<w:vAlign w:val="both"/>') +
				table('<w:bidiVisual/><w:jc w:val="start"/>') +
				table('<w:jc w:val="both"/>') +
				'<w:sectPr/>',
		);
		const tables = loaded.model.blocks.map((block) => (block.type === 'table' ? block : undefined));
		expect(tables[0]).toMatchObject({ alignment: 'right', justification: 'end' });
		expect(at(at(expectTable(at(loaded.model.blocks, 0)).rows, 0), 0).verticalAlign).toBe('both');
		expect(tables[1]).toMatchObject({ alignment: 'right', justification: 'start' });
		expect(tables[2]?.alignment).toBeUndefined();
		expect(loaded.model.warnings.join('\n')).toContain('w:tblPr/w:jc value “both”');
	});
});

describe('typed attribute parsing', () => {
	it('does not treat an invalid on/off value as true', async () => {
		const loaded = await load(
			`<w:p><w:pPr><w:keepNext w:val="banana"/><w:keepLines w:val="off"/><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:b w:val="maybe"/><w:i w:val="1"/><w:strike w:val="on"/></w:rPr><w:t>x</w:t></w:r></w:p><w:sectPr/>`,
		);
		const para = paragraph(loaded);
		expect(para.keepNext).toBeUndefined();
		expect(para.keepLines).toBe(false);
		expect(para.pageBreakBefore).toBe(true);
		expect(para.runs[0]).toMatchObject({ italic: true, strike: true });
		expect(at(para.runs, 0).bold).toBeUndefined();
	});

	it('reads 1in margins and universal-measure spacing as twips', async () => {
		const loaded = await load(
			`<w:p><w:pPr><w:spacing w:before="0.5in"/><w:ind w:left="1in"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="8.5in" w:h="11in"/><w:pgMar w:top="1in" w:right="2.54cm" w:bottom="72pt" w:left="1440" w:header="0.5in"/></w:sectPr>`,
		);
		expect(loaded.model.sections?.[0]).toMatchObject({
			pageWidthTwips: 12240,
			pageHeightTwips: 15840,
			marginTopTwips: 1440,
			marginRightTwips: 1440,
			marginBottomTwips: 1440,
			marginLeftTwips: 1440,
			headerDistanceTwips: 720,
		});
		expect(paragraph(loaded)).toMatchObject({ spacingBeforeTwips: 720, indentLeftTwips: 1440 });
	});

	it('drops invalid enumerations with warnings instead of keeping them', async () => {
		const loaded = await load(
			`<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="bogus"/><w:highlight w:val="mauve"/><w:sz w:val="abc"/></w:rPr><w:t>x</w:t></w:r></w:p>
<w:sectPr><w:type w:val="sideways"/><w:pgNumType w:fmt="roman9"/><w:vAlign w:val="middle"/></w:sectPr>`,
		);
		const run = at(paragraph(loaded).runs, 0);
		expect(run.color).toBe('#112233');
		expect(run.colorTheme).toBeUndefined();
		expect(run.highlight).toBeUndefined();
		expect(run.fontSize).toBeUndefined();
		const section = at(loaded.model.sections, 0);
		expect(section.type).toBe('nextPage');
		expect(section.verticalAlign).toBeUndefined();
		expect(section.pageNumbering?.format).toBeUndefined();
		const warnings = loaded.model.warnings.join('\n');
		for (const bad of ['bogus', 'mauve', 'sideways', 'roman9', 'middle'])
			expect(warnings).toContain(`“${bad}”`);
	});

	it('keeps valid schema enumerations, including tab alignments and leaders', async () => {
		const loaded = await load(
			p(
				'<w:tabs><w:tab w:val="end" w:leader="middleDot" w:pos="4680"/><w:tab w:val="wobble" w:pos="1"/></w:tabs>',
			) +
				`<w:sectPr><w:type w:val="oddPage"/><w:pgNumType w:fmt="upperRoman"/><w:vAlign w:val="both"/></w:sectPr>`,
		);
		expect(paragraph(loaded).tabStops).toEqual([
			{ posTwips: 4680, align: 'end', leader: 'middleDot' },
		]);
		expect(loaded.model.sections![0]).toMatchObject({
			type: 'oddPage',
			verticalAlign: 'both',
			pageNumbering: { format: 'upperRoman' },
		});
	});
});
