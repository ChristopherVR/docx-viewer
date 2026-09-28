import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	applyThemeShade,
	applyThemeTint,
	loadDocx,
	parseColorSchemeMapping,
	parseTheme,
	resolveThemeColorReference,
	resolveThemeColorToken,
} from './index.js';
import { at, expectParagraph } from './test-support/access.js';

const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const themeXml = `<a:theme xmlns:a="${a}"><a:themeElements>
<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="MajorEA"/><a:cs typeface="MajorCS"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="MinorEA"/><a:cs typeface="MinorCS"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;
const settingsXml = `<w:settings xmlns:w="${w}"><w:clrSchemeMapping w:bg1="light1" w:tx1="dark1" w:bg2="light2" w:tx2="dark2"/></w:settings>`;

describe('theme parsing and color resolution', () => {
	it('parses the color scheme and major/minor font scheme', () => {
		const theme = parseTheme(themeXml);
		expect(theme.colors).toMatchObject({
			dk1: '000000',
			lt1: 'FFFFFF',
			dk2: '44546A',
			accent1: '4472C4',
			accent6: '70AD47',
			hlink: '0563C1',
			folHlink: '954F72',
		});
		expect(theme.fonts.major).toMatchObject({
			latin: 'Calibri Light',
			eastAsia: 'MajorEA',
			complexScript: 'MajorCS',
		});
		expect(theme.fonts.minor).toMatchObject({ latin: 'Calibri', eastAsia: 'MinorEA' });
	});

	it('parses the settings.xml color scheme mapping', () => {
		expect(parseColorSchemeMapping(settingsXml)).toEqual({
			bg1: 'lt1',
			tx1: 'dk1',
			bg2: 'lt2',
			tx2: 'dk2',
		});
	});

	it('resolves theme color tokens including logical background/text aliases', () => {
		const theme = parseTheme(themeXml);
		theme.colorMapping = parseColorSchemeMapping(settingsXml);
		expect(resolveThemeColorToken('accent1', theme)).toBe('4472C4');
		expect(resolveThemeColorToken('background1', theme)).toBe('FFFFFF');
		expect(resolveThemeColorToken('text1', theme)).toBe('000000');
		expect(resolveThemeColorToken('hyperlink', theme)).toBe('0563C1');
	});

	it('resolves logical background/text tokens to Word conventional defaults with no mapping', () => {
		const theme = parseTheme(themeXml);
		expect(resolveThemeColorToken('background1', theme)).toBe('FFFFFF');
		expect(resolveThemeColorToken('text1', theme)).toBe('000000');
	});

	it('applies linear-light tint/shade transforms toward white/black', () => {
		expect(applyThemeTint('#4472C4', 1)).toBe('#4472C4');
		expect(applyThemeTint('#4472C4', 0)).toBe('#FFFFFF');
		expect(applyThemeShade('#4472C4', 1)).toBe('#4472C4');
		expect(applyThemeShade('#4472C4', 0)).toBe('#000000');
		const tinted = applyThemeTint('#4472C4', 0.4);
		expect(tinted).not.toBe('#4472C4');
		expect(tinted).toMatch(/^#[0-9A-F]{6}$/);
	});

	it('resolves a full theme color reference (token + tint) end to end', () => {
		const theme = parseTheme(themeXml);
		const resolved = resolveThemeColorReference({ token: 'accent1', tint: 0.4 }, theme);
		expect(resolved).toBe(applyThemeTint('#4472C4', 0.4));
	});

	it('parses word/theme/theme1.xml and settings.xml when loading a package, keeping direct refs unflattened', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:rPr><w:color w:val="000000" w:themeColor="accent1" w:themeTint="66"/></w:rPr><w:t>Themed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file('word/theme/theme1.xml', themeXml);
		zip.file('word/settings.xml', settingsXml);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(loaded.model.theme?.colors.accent1).toBe('4472C4');
		expect(loaded.model.theme?.colorMapping).toEqual({
			bg1: 'lt1',
			tx1: 'dk1',
			bg2: 'lt2',
			tx2: 'dk2',
		});
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(at(paragraph.runs, 0).colorTheme).toMatchObject({ token: 'accent1' });
		expect(at(paragraph.runs, 0).colorTheme?.tint).toBeCloseTo(0x66 / 255, 5);
		expect(at(paragraph.runs, 0).color).toBe('#000000');
		expect(loaded.model.warnings.join(' ')).toContain(
			'Theme colors and fonts resolve for rendering',
		);
		// No-op save preserves the original bytes; the theme reference is never flattened onto the run.
		const original = await zip.generateAsync({ type: 'uint8array' });
		expect(await loaded.save()).toEqual(original);
	});
});
