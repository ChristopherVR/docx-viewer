import { describe, expect, it } from 'vitest';
import {
	createDocument,
	DocxModelValidationError,
	assertValidDocumentModel,
	saveDocx,
	validateDocumentModel,
	type DocumentModel,
	type Paragraph,
	type SectionProperties,
	type TextRun,
} from './index.js';
import { isXsdDateTime } from './validate-issues.js';

const cell = (id: string, extra: object = {}) => ({
	paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text: id }] }],
	...extra,
});

function section(extra: Partial<SectionProperties> = {}): SectionProperties {
	return {
		endsAtBlockId: 'p1',
		type: 'nextPage',
		pageWidthTwips: 12240,
		pageHeightTwips: 15840,
		orientation: 'portrait',
		marginTopTwips: 1440,
		marginRightTwips: 1440,
		marginBottomTwips: 1440,
		marginLeftTwips: 1440,
		columns: { count: 1, equalWidth: true },
		...extra,
	};
}

type Mutate = (model: DocumentModel) => void;
const withRun =
	(patch: Partial<TextRun>): Mutate =>
	(model) => {
		Object.assign((model.blocks[0] as Paragraph).runs[0], patch);
	};
const withParagraph =
	(patch: Partial<Paragraph>): Mutate =>
	(model) => {
		Object.assign(model.blocks[0] as Paragraph, patch);
	};
const withSection =
	(patch: Partial<SectionProperties>): Mutate =>
	(model) => {
		model.sections = [section(patch)];
	};
const withTableCell =
	(patch: object): Mutate =>
	(model) => {
		model.blocks.push({ type: 'table', id: 't', rows: [[cell('c1', patch)]] });
	};
const withTable =
	(patch: object): Mutate =>
	(model) => {
		model.blocks.push({ type: 'table', id: 't', rows: [[cell('c1')]], ...patch });
	};

interface Rule {
	name: string;
	field: string;
	valid: Mutate;
	invalid: Mutate;
}

const rules: Rule[] = [
	{
		name: 'fontSize is whole half-points',
		field: 'fontSize',
		valid: withRun({ fontSize: 10.5 }),
		invalid: withRun({ fontSize: -1 }),
	},
	{
		name: 'fontSize rejects NaN',
		field: 'fontSize',
		valid: withRun({ fontSize: 12 }),
		invalid: withRun({ fontSize: Number.NaN }),
	},
	{
		name: 'fontSize rejects Infinity',
		field: 'fontSize',
		valid: withRun({ fontSize: 12 }),
		invalid: withRun({ fontSize: Infinity }),
	},
	{
		name: 'color is ST_HexColor',
		field: 'color',
		valid: withRun({ color: '#FF00aa' }),
		invalid: withRun({ color: 'rgb(1, 2, 3)' }),
	},
	{
		name: 'color accepts auto',
		field: 'color',
		valid: withRun({ color: 'auto' }),
		invalid: withRun({ color: '#auto' }),
	},
	{
		name: 'shadingFill is ST_HexColor',
		field: 'shadingFill',
		valid: withRun({ shadingFill: 'FFFF00' }),
		invalid: withRun({ shadingFill: 'yellow' }),
	},
	{
		name: 'underlineColor is ST_HexColor',
		field: 'underlineColor',
		valid: withRun({ underlineColor: '#00FF00' }),
		invalid: withRun({ underlineColor: '#0F0' }),
	},
	{
		name: 'underlineStyle is ST_Underline',
		field: 'underlineStyle',
		valid: withRun({ underlineStyle: 'wavyDouble' }),
		invalid: withRun({ underlineStyle: 'squiggle' as never }),
	},
	{
		name: 'colorTheme token is ST_ThemeColor',
		field: 'token',
		valid: withRun({ colorTheme: { token: 'accent1' } }),
		invalid: withRun({ colorTheme: { token: 'accent9' as never } }),
	},
	{
		name: 'colorTheme tint is a fraction',
		field: 'tint',
		valid: withRun({ colorTheme: { token: 'accent1', tint: 0.5 } }),
		invalid: withRun({ colorTheme: { token: 'accent1', tint: 2 } }),
	},
	{
		name: 'shadingThemeFill token is ST_ThemeColor',
		field: 'token',
		valid: withRun({ shadingThemeFill: { token: 'dark1' } }),
		invalid: withRun({ shadingThemeFill: { token: 'nope' as never } }),
	},
	{
		name: 'highlight is ST_HighlightColor',
		field: 'highlight',
		valid: withRun({ highlight: 'darkYellow' }),
		invalid: withRun({ highlight: 'gray25' as never }),
	},
	{
		name: 'break is ST_BrType',
		field: 'break',
		valid: withRun({ break: 'column' }),
		invalid: withRun({ break: 'line' as never }),
	},
	{
		name: 'language is BCP 47',
		field: 'language',
		valid: withRun({ language: 'zh-Hans-CN' }),
		invalid: withRun({ language: 'en--US' }),
	},
	{
		name: 'revision.date is xsd:dateTime',
		field: 'date',
		valid: withRun({
			revision: { kind: 'insert', author: 'A', id: '1', date: '2024-02-29T10:00:00Z' },
		}),
		invalid: withRun({ revision: { kind: 'insert', author: 'A', id: '1', date: 'yesterday' } }),
	},
	{
		name: 'markRevision.date is xsd:dateTime',
		field: 'date',
		valid: withParagraph({
			markRevision: { kind: 'insert', author: 'A', id: '1', date: '2024-01-01T00:00:00+02:00' },
		}),
		invalid: withParagraph({
			markRevision: { kind: 'insert', author: 'A', id: '1', date: '2023-02-29T00:00:00Z' },
		}),
	},
	{
		name: 'comment.date is xsd:dateTime',
		field: 'date',
		valid: (m) => {
			m.comments = [{ id: 'c', author: 'A', text: 'x', date: '2024-01-01T00:00:00Z' }];
		},
		invalid: (m) => {
			m.comments = [{ id: 'c', author: 'A', text: 'x', date: '01/02/2024' }];
		},
	},
	{
		name: 'spacingBeforeTwips is unsigned',
		field: 'spacingBeforeTwips',
		valid: withParagraph({ spacingBeforeTwips: 0 }),
		invalid: withParagraph({ spacingBeforeTwips: -20 }),
	},
	{
		name: 'spacingAfterTwips is an integer',
		field: 'spacingAfterTwips',
		valid: withParagraph({ spacingAfterTwips: 120 }),
		invalid: withParagraph({ spacingAfterTwips: 12.5 }),
	},
	{
		name: 'hangingTwips is unsigned',
		field: 'hangingTwips',
		valid: withParagraph({ hangingTwips: 360 }),
		invalid: withParagraph({ hangingTwips: -360 }),
	},
	{
		name: 'firstLineTwips is unsigned',
		field: 'firstLineTwips',
		valid: withParagraph({ firstLineTwips: 360 }),
		invalid: withParagraph({ firstLineTwips: Number.NaN }),
	},
	{
		name: 'indentLeftTwips is a signed integer',
		field: 'indentLeftTwips',
		valid: withParagraph({ indentLeftTwips: -720 }),
		invalid: withParagraph({ indentLeftTwips: 1.5 }),
	},
	{
		name: 'indentRightTwips is a signed integer',
		field: 'indentRightTwips',
		valid: withParagraph({ indentRightTwips: -1 }),
		invalid: withParagraph({ indentRightTwips: Infinity }),
	},
	{
		name: 'lineSpacingRule is ST_LineSpacingRule',
		field: 'lineSpacingRule',
		valid: withParagraph({ lineSpacingRule: 'atLeast' }),
		invalid: withParagraph({ lineSpacingRule: 'double' as never }),
	},
	{
		name: 'justification is ST_Jc',
		field: 'justification',
		valid: withParagraph({ justification: 'distribute' }),
		invalid: withParagraph({ justification: 'full' as never }),
	},
	{
		name: 'tab stop pos is an integer',
		field: 'posTwips',
		valid: withParagraph({ tabStops: [{ posTwips: 4680, align: 'left' }] }),
		invalid: withParagraph({ tabStops: [{ posTwips: 46.8, align: 'left' }] }),
	},
	{
		name: 'tab stop align is ST_TabJc',
		field: 'align',
		valid: withParagraph({ tabStops: [{ posTwips: 1, align: 'decimal' }] }),
		invalid: withParagraph({ tabStops: [{ posTwips: 1, align: 'middle' as never }] }),
	},
	{
		name: 'table border style is ST_Border',
		field: 'style',
		valid: withTable({ borders: { top: { style: 'double' } } }),
		invalid: withTable({ borders: { top: { style: 'dashy' as never } } }),
	},
	{
		name: 'table border themeColor is ST_ThemeColor',
		field: 'themeColor',
		valid: withTable({ borders: { top: { themeColor: 'accent2' } } }),
		invalid: withTable({ borders: { top: { themeColor: 'accent2x' as never } } }),
	},
	{
		name: 'table border color is ST_HexColor',
		field: 'color',
		valid: withTable({ borders: { top: { color: '#123456' } } }),
		invalid: withTable({ borders: { top: { color: 'red' } } }),
	},
	{
		name: 'table justification is ST_JcTable',
		field: 'justification',
		valid: withTable({ justification: 'end' }),
		invalid: withTable({ justification: 'both' as never }),
	},
	{
		name: 'table width is unsigned',
		field: 'widthTwips',
		valid: withTable({ widthTwips: 9000 }),
		invalid: withTable({ widthTwips: -1 }),
	},
	{
		name: 'cell border style is ST_Border',
		field: 'style',
		valid: withTableCell({ borders: { left: { style: 'single' } } }),
		invalid: withTableCell({ borders: { left: { style: 'x' as never } } }),
	},
	{
		name: 'cell gridSpan is positive',
		field: 'gridSpan',
		valid: withTableCell({ gridSpan: 2 }),
		invalid: withTableCell({ gridSpan: 0 }),
	},
	{
		name: 'cell verticalMerge is ST_Merge',
		field: 'verticalMerge',
		valid: withTableCell({ verticalMerge: 'restart' }),
		invalid: withTableCell({ verticalMerge: 'merge' }),
	},
	{
		name: 'cell verticalAlign is ST_VerticalJc',
		field: 'verticalAlign',
		valid: withTableCell({ verticalAlign: 'center' }),
		invalid: withTableCell({ verticalAlign: 'middle' }),
	},
	{
		name: 'cell shadingFill is ST_HexColor',
		field: 'shadingFill',
		valid: withTableCell({ shadingFill: 'D9D9D9' }),
		invalid: withTableCell({ shadingFill: 'D9D9D' }),
	},
	{
		name: 'cell width is unsigned',
		field: 'widthTwips',
		valid: withTableCell({ widthTwips: 2000 }),
		invalid: withTableCell({ widthTwips: 20.5 }),
	},
	{
		name: 'cell margins are integers',
		field: 'left',
		valid: withTableCell({ margins: { left: 108 } }),
		invalid: withTableCell({ margins: { left: 10.8 } }),
	},
	{
		name: 'pgNumType fmt is ST_NumberFormat',
		field: 'format',
		valid: withSection({ pageNumbering: { format: 'upperRoman', start: 3 } }),
		invalid: withSection({ pageNumbering: { format: 'roman' as never } }),
	},
	{
		name: 'pgNumType start is an integer',
		field: 'start',
		valid: withSection({ pageNumbering: { start: 0 } }),
		invalid: withSection({ pageNumbering: { start: 1.5 } }),
	},
	{
		name: 'section page size is an integer',
		field: 'pageWidthTwips',
		valid: withSection({ pageWidthTwips: 11906 }),
		invalid: withSection({ pageWidthTwips: 11906.5 }),
	},
	{
		name: 'section margins are integers',
		field: 'marginLeftTwips',
		valid: withSection({ marginLeftTwips: 0 }),
		invalid: withSection({ marginLeftTwips: Number.NaN }),
	},
	{
		name: 'section type is ST_SectionMark',
		field: 'type',
		valid: withSection({ type: 'continuous' }),
		invalid: withSection({ type: 'next' as never }),
	},
	{
		name: 'section columns count is bounded',
		field: 'count',
		valid: withSection({ columns: { count: 3, equalWidth: true } }),
		invalid: withSection({ columns: { count: 0, equalWidth: true } }),
	},
	{
		name: 'numbering numFmt is ST_NumberFormat',
		field: 'numFmt',
		valid: (m) => {
			m.numberingCatalog = catalog('lowerLetter');
		},
		invalid: (m) => {
			m.numberingCatalog = catalog('letters' as never);
		},
	},
	{
		name: 'note number format is ST_NumberFormat',
		field: 'footnoteNumFmt',
		valid: (m) => {
			m.footnoteNumFmt = 'chicago';
		},
		invalid: (m) => {
			m.footnoteNumFmt = 'star' as never;
		},
	},
	{
		name: 'page geometry is finite',
		field: 'width',
		valid: (m) => {
			m.page.width = 816;
		},
		invalid: (m) => {
			m.page.width = Number.NaN;
		},
	},
];

function catalog(numFmt: 'lowerLetter') {
	return {
		abstractNums: {
			'1': { id: '1', levels: { 0: { level: 0, start: 1, numFmt, lvlText: '%1.' } } },
		},
		nums: {},
		warnings: [],
	};
}

function base(): DocumentModel {
	const model = createDocument();
	model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text: 'x' }] };
	return model;
}

describe('pre-save model validation', () => {
	it('accepts a plain new document', () => {
		expect(validateDocumentModel(base())).toEqual([]);
		expect(() => assertValidDocumentModel(base())).not.toThrow();
	});

	for (const rule of rules)
		it(`${rule.name}: valid passes, invalid reports ${rule.field}`, () => {
			const valid = base();
			rule.valid(valid);
			expect(validateDocumentModel(valid)).toEqual([]);
			const invalid = base();
			rule.invalid(invalid);
			const issues = validateDocumentModel(invalid);
			expect(issues).toHaveLength(1);
			expect(issues[0].field).toBe(rule.field);
			expect(issues[0].rule).toBeTruthy();
			expect(issues[0].path).toBeTruthy();
		});

	it('reports the path, field, value and rule of each issue', () => {
		const model = base();
		model.blocks.push({
			type: 'table',
			id: 't',
			rows: [[cell('c1'), cell('c2', { shadingFill: 'zz' })]],
		});
		expect(validateDocumentModel(model)).toEqual([
			{
				path: 'blocks[1].rows[0][1]',
				field: 'shadingFill',
				value: 'zz',
				rule: expect.stringContaining('ST_HexColor'),
			},
		]);
	});

	it('walks headers, footers, notes and table paragraphs', () => {
		const model = base();
		const bad: Paragraph = { type: 'paragraph', id: 'bad', runs: [{ text: 'x', fontSize: -2 }] };
		model.sections = [section({ headers: { default: { blocks: [bad] } } })];
		model.footnotes = [{ id: '1', blocks: [{ ...bad, id: 'n' }] }];
		model.blocks.push({ type: 'table', id: 't', rows: [[{ paragraphs: [{ ...bad, id: 'tp' }] }]] });
		const paths = validateDocumentModel(model).map((issue) => issue.path);
		expect(paths).toEqual([
			'blocks[1].rows[0][0].paragraphs[0].runs[0]',
			'sections[0].headers.default.blocks[0].runs[0]',
			'footnotes[0].blocks[0].runs[0]',
		]);
	});

	it('checks xsd:dateTime calendar and clock ranges', () => {
		for (const good of [
			'2024-01-31T09:30:00Z',
			'2024-01-31T09:30:00.123-05:00',
			'2024-01-31T24:00:00',
			'2000-02-29T00:00:00',
		])
			expect(isXsdDateTime(good), good).toBe(true);
		for (const bad of [
			'2024-13-01T00:00:00Z',
			'2024-04-31T00:00:00Z',
			'2100-02-29T00:00:00Z',
			'2024-01-01T25:00:00Z',
			'2024-01-01T10:60:00Z',
			'2024-01-01',
			'2024-01-01 10:00:00',
			'',
		])
			expect(isXsdDateTime(bad), bad).toBe(false);
	});
});

describe('saveDocx validation', () => {
	it('throws a typed error listing every issue and writes nothing', async () => {
		const model = base();
		Object.assign((model.blocks[0] as Paragraph).runs[0], { fontSize: -3, color: 'blue' });
		(model.blocks[0] as Paragraph).spacingAfterTwips = 1.5;
		const error = await saveDocx(model).then(
			() => undefined,
			(caught: unknown) => caught,
		);
		expect(error).toBeInstanceOf(DocxModelValidationError);
		const validation = error as DocxModelValidationError;
		expect(validation.issues.map((issue) => issue.field).sort()).toEqual([
			'color',
			'fontSize',
			'spacingAfterTwips',
		]);
		expect(validation.message).toContain('blocks[0].runs[0].fontSize');
		expect(validation.message).toContain('ST_HexColor');
		expect(validation.message).toContain('got "blue"');
	});

	it('reports schema-invalid section values before any part is written', async () => {
		const model = base();
		model.sections = [
			section({
				pageNumbering: { format: 'bogus' as never },
			}),
		];
		await expect(saveDocx(model)).rejects.toThrow(/sections\[0\]\.pageNumbering\.format/);
	});
});
