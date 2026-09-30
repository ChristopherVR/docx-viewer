import { createDocument, twips } from '@christophervr/docx-core';
import { expect, it } from 'vitest';
import { withBlankHeaderFooter } from './header-footer-commands';
import { withHeaderFooterDistance } from './header-footer-position';

it('changes only one distance in the current section without rounding other measures', () => {
	const model = withBlankHeaderFooter(createDocument(), 'headers', () => 'h');
	const first = model.sections![0]!;
	model.sections = [
		first,
		{
			...first,
			endsAtBlockId: 'second',
			marginLeftTwips: twips(1441),
			footerDistanceTwips: twips(541),
		},
	];
	const original = structuredClone(model);
	const next = withHeaderFooterDistance(model, 1, 'header', 0.75);
	expect(next.sections![0]).toBe(first);
	expect(next.sections![1]).toEqual({ ...original.sections![1], headerDistanceTwips: 1080 });
	expect(next.page).toBe(model.page);
	expect(model).toEqual(original);
	expect(withHeaderFooterDistance(next, 1, 'footer', 0).sections![1]!.footerDistanceTwips).toBe(0);
});
it('keeps imported automatic defaults and rejects invalid or out-of-range input', () => {
	const model = createDocument();
	expect(withHeaderFooterDistance(model, 0, 'header', 0.5)).toBe(model);
	for (const value of [NaN, Infinity, -0.1, 22.001])
		expect(withHeaderFooterDistance(model, 0, 'header', value)).toBe(model);
	expect(withHeaderFooterDistance(model, 2, 'header', 0.75)).toBe(model);
	expect(
		withHeaderFooterDistance(model, 0, 'footer', 0.375).sections![0]!.footerDistanceTwips,
	).toBe(540);
});
