import { describe, expect, it } from 'vitest';
import { alignFragments } from './paragraph-fragments.js';
import type { LayoutFragment } from './result.js';
import { at } from './__tests__/helpers.js';

const frag = (text: string, xPx: number, widthPx: number): LayoutFragment => ({
	text,
	xPx,
	widthPx,
	runIndex: 0,
});
const line = () => [frag('a', 0, 10), frag(' ', 10, 5), frag('b', 15, 10)];

describe('alignFragments', () => {
	it('leaves left-aligned lines untouched', () => {
		expect(alignFragments(line(), 100, false, 'left')).toEqual(line());
	});
	it('offsets right and center alignment by the free width', () => {
		expect(at(alignFragments(line(), 100, false, 'right'), 0).xPx).toBe(75);
		expect(at(alignFragments(line(), 100, false, 'center'), 0).xPx).toBe(37.5);
	});
	it('stretches spaces when justifying, except on the last line', () => {
		const out = alignFragments(line(), 45, false, 'justify');
		expect(at(out, 1).widthPx).toBe(25);
		expect(at(out, 2).xPx).toBe(35);
		expect(alignFragments(line(), 45, true, 'justify')).toEqual(line());
	});
});
