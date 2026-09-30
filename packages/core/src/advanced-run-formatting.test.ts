import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	createDocument,
	loadDocx,
	saveDocx,
	resolveRunFormatting,
	parseRunStyleCatalog,
	parseParagraphStyleCatalog,
	type TextRun,
} from './index.js';
import { halfPoints } from './units.js';
import { parseRunProperties } from './run-properties.js';
import { parseSignedHalfPoints, parseTextScale } from './simple-types.js';
import { parseXml, WORD_NS } from './xml.js';
import { runHasUnknownProperties } from './write-run-validation.js';
import { expectParagraph } from './test-support/access.js';

const rPr = '<w:w w:val="125"/><w:kern w:val="24"/><w:position w:val="-6"/>';
const run = (properties: string) =>
	parseXml(`<w:r xmlns:w="${WORD_NS}"><w:rPr>${properties}</w:rPr><w:t>AV</w:t></w:r>`)
		.documentElement;

async function imported(properties = rPr) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:rPr>${properties}</w:rPr><w:t>AV</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

describe('advanced run formatting preservation', () => {
	it('uses the default paragraph style when no explicit pStyle is present, as Word does', () => {
		const xml = `<w:styles xmlns:w="${WORD_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:rFonts w:ascii="Times New Roman"/><w:sz w:val="32"/><w:position w:val="6"/><w:spacing w:val="40"/><w:kern w:val="24"/></w:rPr></w:style></w:styles>`;
		const context = {
			paragraphCatalog: parseParagraphStyleCatalog(xml),
			runCatalog: parseRunStyleCatalog(xml),
		};
		expect(resolveRunFormatting({ text: 'Implicit Normal' }, context)).toMatchObject({
			fontSize: 16,
			fontFamily: 'Times New Roman',
			positionHalfPoints: 6,
			characterSpacingTwips: 40,
			kerningHalfPoints: 24,
		});
		expect(
			resolveRunFormatting(
				{
					text: 'Direct',
					fontSize: 10,
					positionHalfPoints: halfPoints(0),
					kerningHalfPoints: halfPoints(0),
				},
				context,
			),
		).toMatchObject({ fontSize: 10, positionHalfPoints: 0, kerningHalfPoints: 0 });
	});
	it('parses Word COM scale, kerning threshold and lowered baseline without protecting editable text', async () => {
		const { model } = await imported();
		expect(expectParagraph(model.blocks[0]).runs[0]).toMatchObject({
			textScalePercent: 125,
			kerningHalfPoints: 24,
			positionHalfPoints: -6,
		});
		expect(runHasUnknownProperties(run(rPr))).toBe(false);
	});

	it('accepts schema percentage/universal units and explicit defaults', () => {
		expect(parseTextScale('0125%')).toBe(125);
		expect(parseTextScale('600')).toBe(600);
		expect(parseTextScale('0%')).toBe(0);
		expect(parseSignedHalfPoints('-3pt')).toBe(-6);
		expect(parseSignedHalfPoints('2.5pt')).toBe(5);
		const props = parseXml(
			`<w:rPr xmlns:w="${WORD_NS}"><w:w/><w:kern w:val="0"/><w:position w:val="0"/></w:rPr>`,
		).documentElement;
		expect(parseRunProperties(props)).toMatchObject({
			textScalePercent: 100,
			kerningHalfPoints: 0,
			positionHalfPoints: 0,
		});
	});

	it.each([
		'<w:w w:val="601"/>',
		'<w:w w:val="125.5%"/>',
		'<w:kern w:val="-1"/>',
		'<w:position w:val="unknown"/>',
		'<w:kern/>',
		'<w:position w:val="2" w:custom="keep"/>',
	])('keeps invalid or unmodeled properties protected: %s', (properties) => {
		expect(runHasUnknownProperties(run(properties))).toBe(true);
	});

	it('edits imported text without stripping the advanced settings, then resets and removes them', async () => {
		const { model } = await imported();
		const text = expectParagraph(model.blocks[0]).runs[0]!;
		text.text = 'Edited AV';
		let reopened = await loadDocx(await saveDocx(model));
		expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toMatchObject({
			text: 'Edited AV',
			textScalePercent: 125,
			kerningHalfPoints: 24,
			positionHalfPoints: -6,
		});
		text.textScalePercent = 100;
		text.kerningHalfPoints = halfPoints(0);
		text.positionHalfPoints = halfPoints(0);
		reopened = await loadDocx(await saveDocx(model));
		expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toMatchObject({
			textScalePercent: 100,
			kerningHalfPoints: 0,
			positionHalfPoints: 0,
		});
		delete text.textScalePercent;
		delete text.kerningHalfPoints;
		delete text.positionHalfPoints;
		const zip = await JSZip.loadAsync(await saveDocx(model));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).not.toMatch(/<w:(w|kern|position)[\s/>]/);
	});

	it('writes explicit neutral values on new runs and cancels inherited formatting', async () => {
		const catalog = parseRunStyleCatalog(
			`<w:styles xmlns:w="${WORD_NS}"><w:docDefaults><w:rPrDefault><w:rPr>${rPr}</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
		);
		const text = {
			text: 'Plain',
			textScalePercent: 100,
			kerningHalfPoints: halfPoints(0),
			positionHalfPoints: halfPoints(0),
		};
		expect(resolveRunFormatting({ text: 'Inherited' }, { runCatalog: catalog })).toMatchObject({
			textScalePercent: 125,
			kerningHalfPoints: 24,
			positionHalfPoints: -6,
		});
		expect(resolveRunFormatting(text, { runCatalog: catalog })).toMatchObject({
			textScalePercent: 100,
			kerningHalfPoints: 0,
			positionHalfPoints: 0,
		});
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'new', runs: [text] }];
		const reopened = await loadDocx(await saveDocx(model));
		expect(expectParagraph(reopened.model.blocks[0]).runs[0]).toMatchObject(text);
	});

	it.each([
		{ textScalePercent: 601 },
		{ textScalePercent: 12.5 },
		{ kerningHalfPoints: halfPoints(-1) },
		{ positionHalfPoints: NaN },
	])('refuses invalid model settings before export: %s', async (properties) => {
		const model = createDocument();
		model.blocks = [
			{ type: 'paragraph', id: 'bad', runs: [{ text: 'Invalid', ...properties } as TextRun] },
		];
		await expect(saveDocx(model)).rejects.toThrow('Document model is not valid');
	});
});
