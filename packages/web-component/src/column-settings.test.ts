import { createDocument } from '@christophervr/docx-core';
import { describe, expect, it } from 'vitest';
import {
	changeColumnGap,
	changeColumnWidth,
	columnPreset,
	columnTextWidth,
	equalColumnDraft,
	unequalColumnPreset,
	validColumnDraft,
} from './column-settings';
import { sectionsOf } from './section-commands';
describe('column settings', () => {
	it('creates inverse Left/Right presets with exact twip totals', () => {
		const section = sectionsOf(createDocument())[0]!;
		const left = unequalColumnPreset(section, 'left');
		const right = unequalColumnPreset(section, 'right');
		expect(columnPreset(left)).toBe('left');
		expect(columnPreset(right)).toBe('right');
		expect(left.widths!.map((column) => column.widthTwips)).toEqual(
			[...right.widths!.map((column) => column.widthTwips)].reverse(),
		);
		expect(
			left.widths!.reduce((sum, column) => sum + column.widthTwips + (column.spacingTwips ?? 0), 0),
		).toBe(columnTextWidth(section));
	});
	it('keeps the text width fixed as width and spacing change and rejects too-narrow neighbors', () => {
		const draft = equalColumnDraft(9360, 3, 720);
		changeColumnWidth(draft, 0, 2000);
		expect(draft.widths).toEqual([2000, 3280, 2640]);
		changeColumnGap(draft, 1, 360);
		expect(draft.widths[2]).toBe(3000);
		expect(validColumnDraft(draft, 9360)).toBe(true);
		changeColumnWidth(draft, 0, 5000);
		expect(validColumnDraft(draft, 9360)).toBe(false);
	});
});
