// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';

const mount = () => {
	const ribbon = createRibbon();
	document.body.append(ribbon);
	return ribbon;
};
const tab = (ribbon: HTMLElement, name: string) =>
	ribbon.querySelector<HTMLButtonElement>(`#dve-tab-${name}`)!;

afterEach(() => (document.body.innerHTML = ''));

describe('ribbon collapse', () => {
	it('toggles from the button, a double-click on a tab and Ctrl+F1', () => {
		const ribbon = mount();
		const button = ribbon.querySelector<HTMLButtonElement>('.ribbon-collapse')!;
		button.click();
		expect(ribbon.hasAttribute('data-collapsed')).toBe(true);
		expect(button.getAttribute('aria-pressed')).toBe('true');
		tab(ribbon, 'home').dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(ribbon.hasAttribute('data-collapsed')).toBe(false);
		tab(ribbon, 'home').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, bubbles: true }),
		);
		expect(ribbon.hasAttribute('data-collapsed')).toBe(true);
	});

	it('peeks a panel on tab click and dismisses on Escape, a command or an outside click', () => {
		const ribbon = mount();
		ribbon.querySelector<HTMLButtonElement>('.ribbon-collapse')!.click();
		tab(ribbon, 'insert').click();
		expect(ribbon.hasAttribute('data-peek')).toBe(true);
		tab(ribbon, 'insert').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }),
		);
		expect(ribbon.hasAttribute('data-peek')).toBe(false);
		tab(ribbon, 'insert').click();
		ribbon.querySelector<HTMLButtonElement>('[aria-label="Insert page break"]')!.click();
		expect(ribbon.hasAttribute('data-peek')).toBe(false);
		tab(ribbon, 'view').click();
		document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		expect(ribbon.hasAttribute('data-peek')).toBe(false);
	});

	it('expanding clears any peek', () => {
		const ribbon = mount();
		const button = ribbon.querySelector<HTMLButtonElement>('.ribbon-collapse')!;
		button.click();
		tab(ribbon, 'layout').click();
		button.click();
		expect(ribbon.hasAttribute('data-collapsed')).toBe(false);
		expect(ribbon.hasAttribute('data-peek')).toBe(false);
	});
});

describe('tab KeyTips', () => {
	it('shows a badge per tab and opens the tab for its letter', () => {
		const ribbon = mount();
		ribbon.dispatchEvent(new CustomEvent('dve-keytips'));
		const badges = [...ribbon.querySelectorAll('.dve-keytip')].map((badge) => badge.textContent);
		expect(badges).toEqual(['H', 'N', 'P', 'S', 'R', 'W']);
		ribbon.dispatchEvent(new KeyboardEvent('keydown', { key: 'n' }));
		expect(tab(ribbon, 'insert').getAttribute('aria-selected')).toBe('true');
		expect(ribbon.querySelector('.dve-keytip')).toBeNull();
	});

	it('hides the badges on an unmatched key or Escape without changing tabs', () => {
		const ribbon = mount();
		ribbon.dispatchEvent(new CustomEvent('dve-keytips'));
		ribbon.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(ribbon.querySelector('.dve-keytip')).toBeNull();
		expect(tab(ribbon, 'home').getAttribute('aria-selected')).toBe('true');
	});
});
