// @vitest-environment jsdom
import { createDocument } from 'docx-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createPageSetupDialog } from './page-setup-dialog';
import {
	applyPageSetup,
	presetOf,
	readPageSetup,
	validatePageSetup,
	type PageSetupValues,
} from './page-setup-model';
import { createRibbon } from './ribbon';
import { sectionsOf } from './section-commands';

afterEach(() => (document.body.innerHTML = ''));

const model = () => createDocument();
const values = (): PageSetupValues => readPageSetup(sectionsOf(model())[0]!);

describe('page setup model', () => {
	it('reads Letter with one inch margins', () => {
		expect(values()).toMatchObject({
			topIn: 1,
			bottomIn: 1,
			leftIn: 1,
			rightIn: 1,
			gutterIn: 0,
			orientation: 'portrait',
			widthIn: 8.5,
			heightIn: 11,
		});
		expect(presetOf(values())).toBe('letter');
	});

	it('validates ranges and the room left for text', () => {
		expect(validatePageSetup(values())).toBeNull();
		expect(validatePageSetup({ ...values(), topIn: -1 })).toBe('range');
		expect(validatePageSetup({ ...values(), leftIn: 23 })).toBe('range');
		expect(validatePageSetup({ ...values(), widthIn: 0 })).toBe('range');
		expect(validatePageSetup({ ...values(), leftIn: 4.1, rightIn: 4 })).toBe('width');
		expect(validatePageSetup({ ...values(), topIn: 5, bottomIn: 5.6 })).toBe('height');
		expect(
			validatePageSetup({ ...values(), orientation: 'landscape', topIn: 4.1, bottomIn: 4 }),
		).toBe('height');
	});

	it('applies margins, distances, gutter and paper to the section', () => {
		const next = applyPageSetup(model(), 0, {
			...values(),
			topIn: 0.5,
			leftIn: 1.25,
			gutterIn: 0.3,
			headerIn: 0.4,
			widthIn: 8.27,
			heightIn: 11.69,
		});
		const section = sectionsOf(next)[0]!;
		expect(section).toMatchObject({
			marginTopTwips: 720,
			marginLeftTwips: 1800,
			gutterTwips: 432,
			headerDistanceTwips: 576,
			pageWidthTwips: 11909,
			pageHeightTwips: 16834,
		});
		expect(next.page.marginLeft).toBeCloseTo(1800 / 15);
	});

	it('swaps the paper for landscape and leaves invalid values unapplied', () => {
		const landscape = applyPageSetup(model(), 0, { ...values(), orientation: 'landscape' });
		const section = sectionsOf(landscape)[0]!;
		expect([section.pageWidthTwips, section.pageHeightTwips]).toEqual([15840, 12240]);
		expect(section.orientation).toBe('landscape');
		expect(readPageSetup(section).widthIn).toBe(8.5);
		const base = model();
		expect(applyPageSetup(base, 0, { ...values(), leftIn: 9, rightIn: 9 })).toBe(base);
	});
});

describe('Page setup dialog', () => {
	function open(canEdit = true) {
		const applied: PageSetupValues[] = [];
		const dialog = createPageSetupDialog({
			section: () => sectionsOf(model())[0],
			canEdit: () => canEdit,
			apply: (v) => applied.push(v),
			restoreFocus: vi.fn(),
		});
		document.body.append(dialog.element);
		dialog.open();
		return { dialog, applied, root: dialog.element };
	}
	const field = (root: ParentNode, label: string) =>
		root.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${label}"]`)!;
	const button = (root: ParentNode, text: string) =>
		[...root.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === text)!;
	const type = (el: HTMLInputElement, value: string) => {
		el.value = value;
		el.dispatchEvent(new Event('input'));
	};

	it('shows the section and applies edited values with OK', () => {
		const { applied, root, dialog } = open();
		expect(field(root, 'Top').value).toBe('1');
		expect(field(root, 'Paper size').value).toBe('letter');
		type(field(root, 'Left'), '0.75');
		type(field(root, 'Gutter'), '0.25');
		button(root, 'OK').click();
		expect(applied).toHaveLength(1);
		expect(applied[0]).toMatchObject({ leftIn: 0.75, gutterIn: 0.25, topIn: 1 });
		expect(dialog.isOpen).toBe(false);
	});

	it('follows the paper size list and falls back to Custom size', () => {
		const { root } = open();
		const size = field(root, 'Paper size');
		size.value = 'a4';
		size.dispatchEvent(new Event('change'));
		expect(field(root, 'Width').value).toBe('8.268');
		expect(field(root, 'Height').value).toBe('11.693');
		type(field(root, 'Width'), '6');
		expect(size.value).toBe('custom');
	});

	it('blocks OK and explains margins that leave no room', () => {
		const { root, applied } = open();
		type(field(root, 'Left'), '5');
		type(field(root, 'Right'), '5');
		expect(button(root, 'OK').disabled).toBe(true);
		expect(root.querySelector('.dve-dialog-message')!.textContent).toContain('text width');
		button(root, 'OK').click();
		expect(applied).toHaveLength(0);
		type(field(root, 'Left'), '1');
		type(field(root, 'Right'), '1');
		expect(button(root, 'OK').disabled).toBe(false);
	});

	it('cancels without applying, and does not open when editing is not allowed', () => {
		const { root, applied, dialog } = open();
		button(root, 'Cancel').click();
		expect(applied).toHaveLength(0);
		expect(root.childElementCount).toBe(0);
		expect(dialog.isOpen).toBe(false);
		const locked = open(false);
		expect(locked.dialog.isOpen).toBe(false);
	});

	it("has a launcher on the Layout tab's Page setup group", () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon
			.querySelector<HTMLButtonElement>(
				'[data-panel="Layout"] [data-label="Page setup"] .ribbon-launcher',
			)!
			.click();
		expect(seen).toEqual([{ type: 'formatDialog', kind: 'pageSetup' }]);
	});
});
