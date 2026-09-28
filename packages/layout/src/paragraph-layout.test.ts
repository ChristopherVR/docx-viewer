import { describe, expect, it } from 'vitest';
import { layoutParagraph } from './paragraph-layout.js';
import type { LayoutParagraph, LayoutRun } from './input.js';
import type { TextMeasurer } from './measure.js';

// A measurer with clean, font-agnostic numbers: 10px per character, 20px lines.
const measurer: TextMeasurer = {
	widthOf: (text) => text.length * 10,
	lineHeightOf: () => 20,
};
const noop = () => {};
const lineText = (fragments: { text: string }[]) => fragments.map((f) => f.text).join('');
function paragraph(runs: LayoutRun[], overrides: Partial<LayoutParagraph> = {}): LayoutParagraph {
	return { kind: 'paragraph', id: 'p1', runs, ...overrides };
}

describe('layoutParagraph line breaking', () => {
	it('wraps at word boundaries to fit the available width', () => {
		const result = layoutParagraph(paragraph([{ text: 'aaaa bbbb cccc' }]), 90, measurer, noop);
		expect(result.lines).toHaveLength(2);
		expect(lineText(result.lines[0].fragments)).toBe('aaaa bbbb');
		expect(lineText(result.lines[1].fragments)).toBe('cccc');
	});

	it('honors manual line breaks without wrapping the rest of the width', () => {
		const result = layoutParagraph(paragraph([{ text: 'one\ntwo' }]), 200, measurer, noop);
		expect(result.lines).toHaveLength(2);
		expect(lineText(result.lines[0].fragments)).toBe('one');
		expect(lineText(result.lines[1].fragments)).toBe('two');
	});

	it('reports an explicit page break as a line boundary', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'first', breakAfter: 'page' }, { text: 'second' }]),
			200,
			measurer,
			noop,
		);
		expect(result.lines).toHaveLength(2);
		expect(result.pageBreakAfterLine.has(0)).toBe(true);
		expect(result.columnBreakAfterLine.size).toBe(0);
		expect(lineText(result.lines[0].fragments)).toBe('first');
		expect(lineText(result.lines[1].fragments)).toBe('second');
	});

	it('advances tabs to the next default half-inch stop', () => {
		const result = layoutParagraph(paragraph([{ text: 'A\tB' }]), 200, measurer, noop);
		const [a, tab, b] = result.lines[0].fragments;
		expect(a).toMatchObject({ text: 'A', xPx: 0, widthPx: 10 });
		expect(tab).toMatchObject({ text: '', xPx: 10, widthPx: 38 }); // 48px stop - 10px used
		expect(b).toMatchObject({ text: 'B', xPx: 48, widthPx: 10 });
	});
});

describe('layoutParagraph alignment', () => {
	it('distributes extra space on justified lines but not the paragraph’s last line', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'aa bb cc dd' }], { align: 'justify' }),
			60,
			measurer,
			noop,
		);
		expect(result.lines).toHaveLength(2);
		const spaceFragments = result.lines[0].fragments.filter((f) => f.text === ' ');
		expect(spaceFragments.every((f) => f.widthPx > 10)).toBe(true);
		const lastLineSpaces = result.lines[1].fragments.filter((f) => f.text === ' ');
		expect(lastLineSpaces.every((f) => f.widthPx === 10)).toBe(true);
	});

	it('applies firstLineTwips only to the paragraph’s first line', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'AB\nCD' }], { align: 'right', firstLineTwips: 150 }),
			100,
			measurer,
			noop,
		);
		// The first line's box starts 10px in and still ends at the right edge: 10 + (90 - 20).
		expect(result.lines[0].fragments[0].xPx).toBe(80);
		expect(result.lines[1].fragments[0].xPx).toBe(80); // 100 - 20
	});

	it('applies hangingTwips as a first-line outdent', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'AB\nCD' }], { align: 'right', hangingTwips: 150 }),
			100,
			measurer,
			noop,
		);
		// The first line's box starts 10px out (-10) and still ends at the right edge: -10 + (110 - 20).
		expect(result.lines[0].fragments[0].xPx).toBe(80);
		expect(result.lines[1].fragments[0].xPx).toBe(80); // 100 - 20
	});

	it('offsets left-aligned lines by the left indent and first-line indent', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'AB\nCD' }], { indentLeftTwips: 300, firstLineTwips: 150 }),
			100,
			measurer,
			noop,
		);
		expect(result.lines[0].fragments[0].xPx).toBe(30); // 20 + 10
		expect(result.lines[1].fragments[0].xPx).toBe(20);
	});
});

describe('layoutParagraph line spacing rules', () => {
	it('scales natural height by the auto multiple', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'x' }], { lineSpacingTwips: 360, lineSpacingRule: 'auto' }),
			200,
			measurer,
			noop,
		);
		expect(result.lines[0].heightPx).toBe(30); // 20 * 1.5
	});

	it('uses the exact height regardless of natural line height', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'x' }], { lineSpacingTwips: 300, lineSpacingRule: 'exact' }),
			200,
			measurer,
			noop,
		);
		expect(result.lines[0].heightPx).toBe(20); // 300 twips = 20px, smaller than natural but forced
	});

	it('uses at-least as a floor over the natural height', () => {
		const tall = layoutParagraph(
			paragraph([{ text: 'x' }], { lineSpacingTwips: 450, lineSpacingRule: 'atLeast' }),
			200,
			measurer,
			noop,
		);
		expect(tall.lines[0].heightPx).toBe(30);
		const short = layoutParagraph(
			paragraph([{ text: 'x' }], { lineSpacingTwips: 150, lineSpacingRule: 'atLeast' }),
			200,
			measurer,
			noop,
		);
		expect(short.lines[0].heightPx).toBe(20);
	});
});
