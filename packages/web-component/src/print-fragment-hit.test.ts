// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { resolveFragmentClick, type PrintFragmentHit } from './print-fragment-hit';

afterEach(() => vi.restoreAllMocks());

function fragment(text: string, start: number, left: number, width: number): PrintFragmentHit {
	const element = document.createElement('span');
	element.textContent = text;
	vi.spyOn(element, 'getBoundingClientRect').mockReturnValue({
		left,
		right: left + width,
		width,
	} as DOMRect);
	return {
		element,
		fragment: {
			text,
			sourceStart: start,
			sourceEnd: start + text.length,
			xPx: left,
			widthPx: width,
			runIndex: 0,
		},
	};
}

it('uses glyph widths for a short indented line instead of distributing across its column', () => {
	const hit = fragment('Wi', 7, 100, 24);
	let start = 0;
	vi.spyOn(document, 'createRange').mockReturnValue({
		setStart: (_node: Node, offset: number) => {
			start = offset;
		},
		setEnd: () => {},
		getBoundingClientRect: () =>
			start === 0 ? { left: 100, right: 120, width: 20 } : { left: 120, right: 124, width: 4 },
	} as unknown as Range);
	expect(resolveFragmentClick([hit], 115)).toBe(8);
	expect(resolveFragmentClick([hit], 123)).toBe(9);
	expect(resolveFragmentClick([hit], 400)).toBe(9);
});

it('never positions a cursor inside a surrogate pair or combining grapheme', () => {
	const hit = fragment('😀e\u0301', 0, 0, 40);
	let start = 0;
	const boundaries: number[] = [];
	vi.spyOn(document, 'createRange').mockReturnValue({
		setStart: (_node: Node, offset: number) => {
			start = offset;
			boundaries.push(offset);
		},
		setEnd: (_node: Node, offset: number) => {
			boundaries.push(offset);
		},
		getBoundingClientRect: () => ({ left: start * 10, right: start * 10 + 20, width: 20 }),
	} as unknown as Range);
	expect(resolveFragmentClick([hit], 16)).toBe(2);
	expect(boundaries).toEqual([0, 2, 2, 4]);
});

it('uses fragment boxes around tab gaps and maps synthetic list labels to source start', () => {
	const marker = fragment('12.', 0, 0, 20);
	marker.fragment.sourceEnd = 0;
	const body = fragment('hello', 0, 80, 50);
	expect(resolveFragmentClick([marker, body], 10)).toBe(0);
	expect(resolveFragmentClick([marker, body], 90)).toBe(1);
	expect(resolveFragmentClick([marker, body], 300)).toBe(5);
});

it('leaves older fragment results to the line fallback', () => {
	const hit = fragment('old', 0, 0, 30);
	delete hit.fragment.sourceStart;
	delete hit.fragment.sourceEnd;
	expect(resolveFragmentClick([hit], 10)).toBeUndefined();
});
