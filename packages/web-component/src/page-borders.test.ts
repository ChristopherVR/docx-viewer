// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { type PageBorders, type SectionProperties } from '@christophervr/docx-core';
import { createPageBordersDialog } from './page-borders-dialog';
import { pageBorderBox } from './print-page-borders';

const page = {
	widthPx: 816,
	heightPx: 1056,
	marginTopPx: 96,
	marginRightPx: 96,
	marginBottomPx: 96,
	marginLeftPx: 96,
	pageInSection: 0,
};
const line = { style: 'single' as const, sizeEighthPoints: 8 as never, spacePoints: 24 };

describe('page border drawing', () => {
	it('measures from the page edge by default and from the text margins on request', () => {
		const edge = pageBorderBox(page, { top: line, left: line, bottom: line, right: line })!;
		expect([edge.style.left, edge.style.top, edge.style.width, edge.style.height]).toEqual([
			'32px',
			'32px',
			'752px',
			'992px',
		]);
		expect(edge.style.pointerEvents).toBe('none');
		const text = pageBorderBox(page, {
			top: line,
			left: line,
			bottom: line,
			right: line,
			offsetFrom: 'text',
		})!;
		expect([text.style.left, text.style.top, text.style.width]).toEqual(['64px', '64px', '688px']);
	});

	it('draws only the sides present and honours the display rule and stacking order', () => {
		const one = pageBorderBox(page, { bottom: { ...line, style: 'double' }, zOrder: 'back' })!;
		expect(one.style.borderBottomStyle).toBe('double');
		expect(one.style.borderTopStyle).toBe('');
		expect(one.style.zIndex).toBe('0');
		expect(
			pageBorderBox({ ...page, pageInSection: 1 }, { top: line, display: 'firstPage' }),
		).toBeNull();
		expect(pageBorderBox(page, { top: line, display: 'notFirstPage' })).toBeNull();
		expect(pageBorderBox(page, { top: { style: 'none' } })).toBeNull();
	});
});

function dialogFor(section: SectionProperties | undefined, canEdit = true) {
	const applied: Array<PageBorders | undefined> = [];
	const dialog = createPageBordersDialog({
		section: () => section,
		canEdit: () => canEdit,
		apply: (borders) => applied.push(borders),
		restoreFocus: () => {},
	});
	document.body.append(dialog.element);
	const field = (name: string) =>
		dialog.element.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${name}"]`)!;
	const button = (text: string) =>
		[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === text)!;
	return { dialog, applied, field, button };
}
const section = () => ({}) as unknown as SectionProperties;

describe('Page Borders dialog', () => {
	it('applies a box with the chosen pen and measure, and None removes it', () => {
		const s = dialogFor(section());
		s.dialog.open();
		expect(s.field('Setting').value).toBe('none');
		s.field('Setting').value = 'box';
		s.field('Setting').dispatchEvent(new Event('change'));
		s.field('Style').value = 'double';
		s.field('Width').value = '12';
		s.field('Measure from').value = 'text';
		s.field('Distance (points)').value = '10';
		s.field('Show on').value = 'firstPage';
		s.button('OK').click();
		expect(s.applied[0]).toMatchObject({
			top: { style: 'double', sizeEighthPoints: 12, spacePoints: 10 },
			right: { style: 'double' },
			offsetFrom: 'text',
			display: 'firstPage',
		});
		const existing = { ...section(), pageBorders: s.applied[0]! };
		const again = dialogFor(existing);
		again.dialog.open();
		expect(again.field('Setting').value).toBe('box');
		expect(again.field('Measure from').value).toBe('text');
		again.field('Setting').value = 'none';
		again.field('Setting').dispatchEvent(new Event('change'));
		again.button('OK').click();
		expect(again.applied).toEqual([undefined]);
	});

	it('requires a side and a whole distance, and does not open when editing is not allowed', () => {
		const s = dialogFor(section());
		s.dialog.open();
		s.field('Setting').value = 'custom';
		s.field('Setting').dispatchEvent(new Event('change'));
		expect(s.button('OK').disabled).toBe(true);
		s.field('Top').checked = true;
		s.field('Top').dispatchEvent(new Event('change'));
		expect(s.button('OK').disabled).toBe(false);
		s.field('Distance (points)').value = '40';
		s.field('Distance (points)').dispatchEvent(new Event('input'));
		expect(s.button('OK').disabled).toBe(true);
		const blocked = dialogFor(section(), false);
		blocked.dialog.open();
		expect(blocked.dialog.isOpen).toBe(false);
	});
});
