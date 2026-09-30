// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createZoomDialog, parseZoomPercent } from './zoom-dialog';

afterEach(() => (document.body.innerHTML = ''));

function setup(percent = 100) {
	const host = {
		percent: () => percent,
		setPercent: vi.fn(),
		fit: vi.fn(),
		restoreFocus: vi.fn(),
	};
	const dialog = createZoomDialog(host);
	document.body.append(dialog.element);
	return { dialog, host };
}
const radio = (root: HTMLElement, label: string) =>
	root.querySelector<HTMLInputElement>(`input[type="radio"][aria-label="${label}"]`)!;
const button = (root: HTMLElement, label: string) =>
	[...root.querySelectorAll('button')].find((b) => b.textContent === label)!;

describe('parseZoomPercent', () => {
	it('accepts numbers with an optional percent sign and clamps to 10-500', () => {
		expect(parseZoomPercent('75')).toBe(75);
		expect(parseZoomPercent(' 120% ')).toBe(120);
		expect(parseZoomPercent('5')).toBe(10);
		expect(parseZoomPercent('900')).toBe(500);
		expect(parseZoomPercent('')).toBeNull();
		expect(parseZoomPercent('abc')).toBeNull();
	});
});

describe('Zoom dialog', () => {
	it('opens on the current zoom and applies a preset', () => {
		const { dialog, host } = setup(100);
		dialog.open();
		expect(radio(dialog.element, '100%').checked).toBe(true);
		radio(dialog.element, '200%').click();
		button(dialog.element, 'OK').click();
		expect(host.setPercent).toHaveBeenCalledWith(200);
		expect(dialog.isOpen).toBe(false);
		expect(host.restoreFocus).toHaveBeenCalled();
	});

	it('selects Percent for an unusual zoom and applies a typed value', () => {
		const { dialog, host } = setup(130);
		dialog.open();
		expect(radio(dialog.element, 'Percent:').checked).toBe(true);
		const field = dialog.element.querySelector<HTMLInputElement>('input[type="number"]')!;
		expect(field.value).toBe('130');
		field.value = '85';
		field.dispatchEvent(new Event('input', { bubbles: true }));
		button(dialog.element, 'OK').click();
		expect(host.setPercent).toHaveBeenCalledWith(85);
	});

	it('fits the page width or whole page, and Cancel changes nothing', () => {
		const { dialog, host } = setup();
		dialog.open();
		radio(dialog.element, 'Whole page').click();
		button(dialog.element, 'OK').click();
		expect(host.fit).toHaveBeenCalledWith('page');
		dialog.open();
		radio(dialog.element, '75%').click();
		button(dialog.element, 'Cancel').click();
		expect(host.setPercent).not.toHaveBeenCalled();
	});
});
