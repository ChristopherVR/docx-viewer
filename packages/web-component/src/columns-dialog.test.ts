// @vitest-environment jsdom
import { createDocument, twips } from '@christophervr/docx-core';
import { describe, expect, it, vi } from 'vitest';
import { createColumnsDialog } from './columns-dialog';
import { sectionsOf } from './section-commands';

describe('More Columns', () => {
	it('validates available width and applies count, spacing and separator', () => {
		const section = sectionsOf(createDocument())[0]!;
		const apply = vi.fn();
		const dialog = createColumnsDialog({
			section: () => section,
			canEdit: () => true,
			apply,
			restoreFocus: () => {},
		});
		dialog.open();
		const input = (label: string) =>
			dialog.element.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
		const count = input('Number of columns');
		const gap = input('Column spacing (inches)');
		const ok = [...dialog.element.querySelectorAll('button')].find(
			(button) => button.textContent === 'OK',
		)!;
		count.value = '3';
		gap.value = '5';
		gap.dispatchEvent(new Event('input'));
		expect(ok.disabled).toBe(true);
		ok.click();
		expect(apply).not.toHaveBeenCalled();
		gap.value = '0.25';
		gap.dispatchEvent(new Event('input'));
		input('Line between').checked = true;
		expect(ok.disabled).toBe(false);
		ok.click();
		expect(apply).toHaveBeenCalledWith({
			count: 3,
			equalWidth: true,
			spacingTwips: 360,
			separator: true,
		});
	});
	it('preserves imported unequal widths until explicitly converted', () => {
		const section = sectionsOf(createDocument())[0]!;
		section.columns = {
			count: 2,
			equalWidth: false,
			widths: [{ widthTwips: twips(2500) }, { widthTwips: twips(6140) }],
			spacingTwips: twips(720),
		};
		const apply = vi.fn();
		const dialog = createColumnsDialog({
			section: () => section,
			canEdit: () => true,
			apply,
			restoreFocus: () => {},
		});
		dialog.open();
		const count = dialog.element.querySelector<HTMLInputElement>(
			'[aria-label="Number of columns"]',
		)!;
		expect(count.disabled).toBe(false);
		const ok = [...dialog.element.querySelectorAll('button')].find(
			(button) => button.textContent === 'OK',
		)!;
		ok.click();
		expect(apply).toHaveBeenCalledWith({
			...section.columns,
			separator: false,
			widths: [{ widthTwips: 2500, spacingTwips: 720 }, { widthTwips: 6140 }],
		});
		dialog.open();
		const equal = dialog.element.querySelector<HTMLInputElement>(
			'[aria-label="Equal column width"]',
		)!;
		equal.checked = true;
		equal.dispatchEvent(new Event('change'));
		expect(count.disabled).toBe(false);
		ok.click();
		expect(apply).toHaveBeenLastCalledWith({
			count: 2,
			equalWidth: true,
			spacingTwips: 720,
			separator: false,
		});
	});
});
