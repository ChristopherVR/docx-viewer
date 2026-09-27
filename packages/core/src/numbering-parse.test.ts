import { describe, expect, it } from 'vitest';
import { parseNumberingCatalog, resolveNumberingLevel } from './numbering-parse.js';

const WORD_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function numberingXml(): string {
	return `<?xml version="1.0"?><w:numbering ${WORD_NS}>
		<w:abstractNum w:abstractNumId="0">
			<w:lvl w:ilvl="0">
				<w:start w:val="1"/>
				<w:numFmt w:val="decimal"/>
				<w:lvlText w:val="%1."/>
				<w:lvlJc w:val="left"/>
				<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
				<w:suff w:val="tab"/>
			</w:lvl>
			<w:lvl w:ilvl="1">
				<w:start w:val="1"/>
				<w:numFmt w:val="lowerLetter"/>
				<w:lvlText w:val="%2)"/>
				<w:lvlRestart w:val="0"/>
			</w:lvl>
		</w:abstractNum>
		<w:abstractNum w:abstractNumId="1">
			<w:numStyleLink w:val="ListParagraph"/>
			<w:lvl w:ilvl="0">
				<w:start w:val="1"/>
				<w:numFmt w:val="bullet"/>
				<w:lvlText w:val=""/>
			</w:lvl>
		</w:abstractNum>
		<w:num w:numId="1">
			<w:abstractNumId w:val="0"/>
		</w:num>
		<w:num w:numId="2">
			<w:abstractNumId w:val="0"/>
			<w:lvlOverride w:ilvl="0">
				<w:startOverride w:val="5"/>
			</w:lvlOverride>
		</w:num>
	</w:numbering>`;
}

describe('numbering-parse', () => {
	it('parses abstract numbering levels, indentation and suffix', () => {
		const catalog = parseNumberingCatalog(numberingXml());
		const level0 = catalog.abstractNums['0'].levels[0];
		expect(level0).toMatchObject({
			level: 0,
			start: 1,
			numFmt: 'decimal',
			lvlText: '%1.',
			lvlJc: 'left',
			indentLeftTwips: 720,
			hangingTwips: 360,
			suffix: 'tab',
		});
		expect(catalog.abstractNums['0'].levels[1]).toMatchObject({
			numFmt: 'lowerLetter',
			lvlRestart: 0,
			suffix: 'tab',
		});
	});

	it('flags style-linked numbering as an unresolved warning', () => {
		const catalog = parseNumberingCatalog(numberingXml());
		expect(catalog.warnings.some((warning) => warning.includes('style-linked'))).toBe(true);
	});

	it('parses num definitions and lvlOverride startOverride', () => {
		const catalog = parseNumberingCatalog(numberingXml());
		expect(catalog.nums['1']).toMatchObject({ id: '1', abstractNumId: '0' });
		expect(catalog.nums['2'].levelOverrides?.[0]).toMatchObject({ startOverride: 5 });
	});

	it('resolves the effective level definition applying startOverride', () => {
		const catalog = parseNumberingCatalog(numberingXml());
		expect(resolveNumberingLevel(catalog, '1', 0)).toMatchObject({ start: 1, numFmt: 'decimal' });
		expect(resolveNumberingLevel(catalog, '2', 0)).toMatchObject({ start: 5, numFmt: 'decimal' });
		expect(resolveNumberingLevel(catalog, 'missing', 0)).toBeUndefined();
	});
});
