import { twips, type SectionColumns, type SectionProperties } from '@christophervr/docx-core';

export const columnTextWidth = (section: SectionProperties): number =>
	section.pageWidthTwips -
	section.marginLeftTwips -
	section.marginRightTwips -
	(section.gutterTwips ?? 0);

/** Word's Left/Right presets reserve one third of the text width for the narrow column. */
export function unequalColumnPreset(
	section: SectionProperties,
	preset: 'left' | 'right',
): SectionColumns {
	const gap = 720;
	const available = columnTextWidth(section) - gap;
	const narrow = Math.round(available / 3);
	const wide = available - narrow;
	return {
		count: 2,
		equalWidth: false,
		spacingTwips: twips(gap),
		widths: [
			{ widthTwips: twips(preset === 'left' ? narrow : wide), spacingTwips: twips(gap) },
			{ widthTwips: twips(preset === 'left' ? wide : narrow) },
		],
	};
}

export function columnPreset(columns: SectionColumns): string {
	if (columns.equalWidth) return String(columns.count);
	const [a, b] = columns.widths ?? [];
	if (columns.count === 2 && a && b) {
		if (Math.abs(a.widthTwips * 2 - b.widthTwips) <= 2) return 'left';
		if (Math.abs(b.widthTwips * 2 - a.widthTwips) <= 2) return 'right';
	}
	return '';
}

export interface ColumnDraft {
	widths: number[];
	gaps: number[];
}
export function equalColumnDraft(total: number, count: number, gap: number): ColumnDraft {
	const width = Math.floor((total - gap * (count - 1)) / count);
	return {
		widths: Array.from({ length: count }, (_, i) =>
			i === count - 1 ? total - gap * (count - 1) - width * (count - 1) : width,
		),
		gaps: Array.from({ length: count }, (_, i) => (i === count - 1 ? 0 : gap)),
	};
}
export function readColumnDraft(section: SectionProperties): ColumnDraft {
	const fallback = equalColumnDraft(
		columnTextWidth(section),
		section.columns.count,
		section.columns.spacingTwips ?? 720,
	);
	return !section.columns.equalWidth && section.columns.widths?.length === section.columns.count
		? {
				widths: section.columns.widths.map((column) => column.widthTwips),
				gaps: section.columns.widths.map((column, i) =>
					i === section.columns.count - 1
						? 0
						: (column.spacingTwips ?? section.columns.spacingTwips ?? 720),
				),
			}
		: fallback;
}
/** Keep the overall text width fixed by taking the difference from the adjacent column. */
export function changeColumnWidth(draft: ColumnDraft, index: number, width: number): void {
	const delta = width - draft.widths[index]!;
	draft.widths[index] = width;
	const neighbor = index < draft.widths.length - 1 ? index + 1 : index - 1;
	if (neighbor >= 0) draft.widths[neighbor] = draft.widths[neighbor]! - delta;
}
export function changeColumnGap(draft: ColumnDraft, index: number, gap: number): void {
	const delta = gap - draft.gaps[index]!;
	draft.gaps[index] = gap;
	if (index + 1 < draft.widths.length) draft.widths[index + 1] = draft.widths[index + 1]! - delta;
}
export function validColumnDraft(draft: ColumnDraft, total: number): boolean {
	return (
		draft.widths.every((width) => Number.isSafeInteger(width) && width >= 720) &&
		draft.gaps.every((gap) => Number.isSafeInteger(gap) && gap >= 0) &&
		Math.abs(draft.widths.reduce((sum, value, i) => sum + value + draft.gaps[i]!, 0) - total) <= 1
	);
}
