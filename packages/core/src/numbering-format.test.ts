import { describe, expect, it } from 'vitest';
import type { DocumentModel, Paragraph } from './model.js';
import type { NumberingCatalog } from './numbering-model.js';
import {
	computeListLabels,
	formatListNumber,
	resolveParagraphNumbering,
} from './numbering-format.js';
import { cardinalWords, letterLabel, ordinalWords, romanNumeral } from './numbering-text.js';
import { at, must } from './test-support/access.js';

function page() {
	return {
		width: 816,
		height: 1056,
		marginTop: 96,
		marginRight: 96,
		marginBottom: 96,
		marginLeft: 96,
	};
}
function paragraph(id: string, numId?: number, level = 0): Paragraph {
	return {
		type: 'paragraph',
		id,
		runs: [{ text: id }],
		...(numId !== undefined ? { numbering: { numId, level } } : {}),
	};
}
function catalogWith(overrides: Partial<NumberingCatalog> = {}): NumberingCatalog {
	return {
		abstractNums: {
			'0': {
				id: '0',
				levels: {
					0: { level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', suffix: 'tab' },
					1: { level: 1, start: 1, numFmt: 'lowerLetter', lvlText: '%2)', suffix: 'tab' },
				},
			},
			'9': {
				id: '9',
				levels: { 0: { level: 0, start: 1, numFmt: 'bullet', lvlText: '•', suffix: 'tab' } },
			},
		},
		nums: { '1': { id: '1', abstractNumId: '0' }, '2': { id: '2', abstractNumId: '9' } },
		warnings: [],
		...overrides,
	};
}

describe('numbering-text formatters', () => {
	it('renders roman numerals and repeating letters', () => {
		expect(romanNumeral(1994)).toBe('MCMXCIV');
		expect(letterLabel(1)).toBe('a');
		expect(letterLabel(26)).toBe('z');
		expect(letterLabel(27)).toBe('aa');
		expect(letterLabel(52)).toBe('zz');
	});

	it('renders cardinal and ordinal words', () => {
		expect(cardinalWords(0)).toBe('zero');
		expect(cardinalWords(21)).toBe('twenty-one');
		expect(cardinalWords(101)).toBe('one hundred one');
		expect(ordinalWords(1)).toBe('first');
		expect(ordinalWords(21)).toBe('twenty-first');
		expect(ordinalWords(12)).toBe('twelfth');
	});

	it('formats list numbers per Word numFmt token', () => {
		expect(formatListNumber('decimalZero', 3)).toBe('03');
		expect(formatListNumber('upperRoman', 9)).toBe('IX');
		expect(formatListNumber('upperLetter', 2)).toBe('B');
		expect(formatListNumber('ordinal', 2)).toBe('2nd');
		expect(formatListNumber('bullet', 1)).toBe('');
		expect(formatListNumber('chicago', 4)).toBe('4');
	});
});

describe('computeListLabels', () => {
	it('numbers a simple list in document order and restarts a new numId at 1', () => {
		const model: DocumentModel = {
			blocks: [paragraph('a', 1, 0), paragraph('b', 1, 0), paragraph('c', 1, 0)],
			page: page(),
			warnings: [],
			numberingCatalog: catalogWith(),
		};
		const labels = computeListLabels(model);
		expect(labels.get('a')?.text).toBe('1.');
		expect(labels.get('b')?.text).toBe('2.');
		expect(labels.get('c')?.text).toBe('3.');
	});

	it('tracks independent multilevel counters and restarts deeper levels by default', () => {
		const model: DocumentModel = {
			blocks: [
				paragraph('a', 1, 0),
				paragraph('a1', 1, 1),
				paragraph('a2', 1, 1),
				paragraph('b', 1, 0),
				paragraph('b1', 1, 1),
			],
			page: page(),
			warnings: [],
			numberingCatalog: catalogWith(),
		};
		const labels = computeListLabels(model);
		expect(labels.get('a')?.text).toBe('1.');
		expect(labels.get('a1')?.text).toBe('a)');
		expect(labels.get('a2')?.text).toBe('b)');
		expect(labels.get('b')?.text).toBe('2.');
		// Level 1 restarts to "a)" because level 0 advanced.
		expect(labels.get('b1')?.text).toBe('a)');
	});

	it('parses lvlRestart for round-tripping without affecting the simplified default restart', () => {
		const catalog = catalogWith({
			abstractNums: {
				'0': {
					id: '0',
					levels: {
						0: { level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', suffix: 'tab' },
						1: {
							level: 1,
							start: 1,
							numFmt: 'decimal',
							lvlText: '%2.',
							suffix: 'tab',
							lvlRestart: 5,
						},
					},
				},
			},
			nums: { '1': { id: '1', abstractNumId: '0' } },
		});
		const model: DocumentModel = {
			blocks: [
				paragraph('a', 1, 0),
				paragraph('a1', 1, 1),
				paragraph('b', 1, 0),
				paragraph('b1', 1, 1),
			],
			page: page(),
			warnings: [],
			numberingCatalog: catalog,
		};
		const labels = computeListLabels(model);
		expect(labels.get('a1')?.text).toBe('1.');
		// The simplified counter always restarts a deeper level on any shallower-level change.
		expect(labels.get('b1')?.text).toBe('1.');
		expect(at(must(catalog.abstractNums['0'], 'abstract numbering 0').levels, 1).lvlRestart).toBe(
			5,
		);
	});

	it('renders bullet levels using the literal lvlText glyph without counting', () => {
		const model: DocumentModel = {
			blocks: [paragraph('a', 2, 0), paragraph('b', 2, 0)],
			page: page(),
			warnings: [],
			numberingCatalog: catalogWith(),
		};
		const labels = computeListLabels(model);
		expect(labels.get('a')?.text).toBe('•');
		expect(labels.get('b')?.text).toBe('•');
	});

	it('forces decimal placeholders when the level uses legal numbering', () => {
		const catalog = catalogWith({
			abstractNums: {
				'0': {
					id: '0',
					levels: {
						0: { level: 0, start: 1, numFmt: 'upperRoman', lvlText: '%1.', suffix: 'tab' },
						1: {
							level: 1,
							start: 1,
							numFmt: 'lowerLetter',
							lvlText: '%1.%2.',
							suffix: 'tab',
							isLgl: true,
						},
					},
				},
			},
			nums: { '1': { id: '1', abstractNumId: '0' } },
		});
		const model: DocumentModel = {
			blocks: [paragraph('a', 1, 0), paragraph('a1', 1, 1)],
			page: page(),
			warnings: [],
			numberingCatalog: catalog,
		};
		const labels = computeListLabels(model);
		expect(labels.get('a')?.text).toBe('I.');
		expect(labels.get('a1')?.text).toBe('1.1.');
	});

	it('resolves numbering inherited from a paragraph style and blocks it via an explicit numId=0 override', () => {
		const model: DocumentModel = {
			blocks: [
				{ type: 'paragraph', id: 'a', runs: [{ text: 'a' }], style: 'ListParagraph' },
				{
					type: 'paragraph',
					id: 'b',
					runs: [{ text: 'b' }],
					style: 'ListParagraph',
					numbering: { numId: 0, level: 0 },
				},
			],
			page: page(),
			warnings: [],
			numberingCatalog: catalogWith(),
			paragraphStyles: {
				docDefaults: {},
				styles: {
					ListParagraph: { id: 'ListParagraph', formatting: {}, numbering: { numId: 1, level: 0 } },
				},
				warnings: [],
			},
		};
		expect(resolveParagraphNumbering(model.blocks[0] as Paragraph, model.paragraphStyles)).toEqual({
			numId: 1,
			level: 0,
		});
		const labels = computeListLabels(model);
		expect(labels.get('a')?.text).toBe('1.');
		expect(labels.has('b')).toBe(false);
	});

	it('flattens table cell paragraphs into document order', () => {
		const model: DocumentModel = {
			blocks: [
				{
					type: 'table',
					id: 't',
					rows: [
						[
							{ paragraphs: [paragraph('cell-a', 1, 0)] },
							{ paragraphs: [paragraph('cell-b', 1, 0)] },
						],
					],
				},
				paragraph('after', 1, 0),
			],
			page: page(),
			warnings: [],
			numberingCatalog: catalogWith(),
		};
		const labels = computeListLabels(model);
		expect(labels.get('cell-a')?.text).toBe('1.');
		expect(labels.get('cell-b')?.text).toBe('2.');
		expect(labels.get('after')?.text).toBe('3.');
	});
});
