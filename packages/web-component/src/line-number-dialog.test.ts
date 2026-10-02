// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { createLineNumberDialog, type LineNumberSettings } from './line-number-dialog';
import { sectionsOf } from './section-commands';
import { createDocument } from 'docx-core';

function setup(settings?: LineNumberSettings) {
	const section = sectionsOf(createDocument())[0]!;
	if (settings) section.lineNumberSettings = settings;
	const apply = vi.fn();
	const dialog = createLineNumberDialog({
		section: () => section,
		canEdit: () => true,
		apply,
		restoreFocus: vi.fn(),
	});
	const input = (label: string) =>
		dialog.element.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
	const ok = () =>
		[...dialog.element.querySelectorAll('button')].find((button) => button.textContent === 'OK')!;
	const change = (label: string, value: string | boolean) => {
		const control = input(label);
		if (typeof value === 'boolean') control.checked = value;
		else control.value = value;
		control.dispatchEvent(new Event('input'));
	};
	dialog.open();
	return { dialog, input, apply, ok, change };
}
describe('Line Numbers options', () => {
	it('reads imported values, validates whole numbers and applies measurements in twips', () => {
		const { dialog, input, apply, ok, change } = setup({
			start: 3,
			countBy: 5,
			restart: 'newSection',
		});
		expect(input('Start at').value).toBe('3');
		expect(input('Distance from text (inches)').disabled).toBe(true);
		change('Count by', '1.5');
		expect(ok().disabled).toBe(true);
		ok().click();
		expect(apply).not.toHaveBeenCalled();
		change('Count by', '2');
		change('Automatic distance from text', false);
		change('Distance from text (inches)', '0.25');
		ok().click();
		expect(apply).toHaveBeenCalledWith({
			start: 3,
			countBy: 2,
			restart: 'newSection',
			distanceTwips: 360,
		});
		expect(dialog.isOpen).toBe(false);
	});
	it('can turn numbering off and cancel edits without applying anything', () => {
		const { dialog, apply, change, ok } = setup({ start: 1, countBy: 1, restart: 'continuous' });
		change('Add line numbering', false);
		ok().click();
		expect(apply).toHaveBeenCalledWith(undefined);
		dialog.open();
		change('Start at', '10');
		dialog.close();
		expect(apply).toHaveBeenCalledTimes(1);
	});
});
