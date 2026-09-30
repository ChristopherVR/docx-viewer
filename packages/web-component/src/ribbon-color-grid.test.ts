// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { openColorGridPopover, themeShades } from './ribbon-color-grid';

describe('colour grid', () => {
	it('derives Word-style variants: white darkens, black lightens, colours lighten then darken', () => {
		expect(themeShades('#ffffff')[4]).toBe('#808080');
		expect(themeShades('#000000')[0]).toBe('#808080');
		const [l80, , , d25, d50] = themeShades('#4472c4');
		expect(l80).toBe('#dae3f3');
		expect(d25).toBe('#335693');
		expect(d50).toBe('#223962');
	});

	it('lists 60 theme swatches and 10 standard ones, and reports the chosen colour', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const anchor = document.createElement('button');
		host.append(anchor);
		const chosen: string[] = [];
		openColorGridPopover(anchor, (value) => chosen.push(value), { noneLabel: 'No Color' });
		const pop = document.querySelector('.color-grid')!;
		expect(pop.querySelectorAll('.color-grid-cells')[0]!.children).toHaveLength(60);
		expect(pop.querySelectorAll('.color-grid-cells')[1]!.children).toHaveLength(10);
		pop.querySelector<HTMLButtonElement>('[aria-label="Purple"]')!.click();
		expect(chosen).toEqual(['#7030a0']);
		expect(document.querySelector('.color-grid')).toBeNull();
		openColorGridPopover(anchor, (value) => chosen.push(value), { noneLabel: 'No Color' });
		document.querySelector<HTMLButtonElement>('.color-grid-command')!.click();
		expect(chosen.at(-1)).toBe('none');
		host.remove();
	});
});
