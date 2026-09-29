import { createDocument, twips } from '@christophervr/docx-core';
import { describe, expect, it } from 'vitest';
import { pageSetupChange } from './page-setup-change';
import { PAGE_SIZES, pageSizeOf, setPageSize } from './page-size';
import { sectionsOf, setMargins } from './section-commands';

const model = () => createDocument();

describe('page size', () => {
	it('sets portrait width and height from the preset', () => {
		const next = setPageSize(model(), 0, 'a4');
		const section = sectionsOf(next)[0]!;
		expect([section.pageWidthTwips, section.pageHeightTwips]).toEqual([11906, 16838]);
		expect(next.page.width).toBeCloseTo(11906 / 15, 0);
	});

	it('keeps a landscape section landscape', () => {
		const landscape = pageSetupChange(
			model(),
			0,
			sectionsOf(model())[0]!,
			'orientation',
			'landscape',
		);
		const next = setPageSize(landscape, 0, 'legal');
		const section = sectionsOf(next)[0]!;
		expect([section.pageWidthTwips, section.pageHeightTwips]).toEqual([20160, 12240]);
		expect(section.orientation).toBe('landscape');
	});

	it('ignores an unknown preset and recognises sizes in either orientation', () => {
		const base = model();
		expect(setPageSize(base, 0, 'nope')).toBe(base);
		const section = sectionsOf(setPageSize(base, 0, 'a5'))[0]!;
		expect(pageSizeOf(section)).toBe('a5');
		expect(
			pageSizeOf({ ...section, pageWidthTwips: twips(8391 + 500), pageHeightTwips: twips(11906) }),
		).toBeNull();
		expect(
			pageSizeOf({ ...section, pageWidthTwips: twips(11906), pageHeightTwips: twips(8391) }),
		).toBe('a5');
		expect(Object.keys(PAGE_SIZES)).toContain('letter');
	});

	it('is reachable through the Layout-tab dispatcher', () => {
		const base = model();
		const next = pageSetupChange(base, 0, sectionsOf(base)[0]!, 'size', 'a3');
		expect(sectionsOf(next)[0]!.pageWidthTwips).toBe(16838);
	});
});

describe('moderate margins', () => {
	it('uses one inch top and bottom and 0.75 inch left and right, as Word does', () => {
		const section = sectionsOf(setMargins(model(), 0, 'moderate'))[0]!;
		expect([
			section.marginTopTwips,
			section.marginBottomTwips,
			section.marginLeftTwips,
			section.marginRightTwips,
		]).toEqual([1440, 1440, 1080, 1080]);
	});
});
