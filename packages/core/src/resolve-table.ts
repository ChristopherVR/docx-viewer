// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TableLook, TableStyleCatalog, TableStyleConditionalFormatting } from './table-model.js';

function styleChain(styleId: string | undefined, catalog: TableStyleCatalog | undefined) {
	if (!styleId || !catalog) return [];
	const chain: NonNullable<TableStyleCatalog['styles'][string]>[] = [];
	const visited = new Set<string>();
	let current: string | undefined = styleId;
	while (current && catalog.styles[current] && !visited.has(current)) {
		visited.add(current);
		chain.push(catalog.styles[current]);
		current = catalog.styles[current].basedOn;
	}
	return chain.reverse();
}

/** Which conditional regions apply to a cell, in increasing precedence (last wins), per `tblLook`. */
function applicableRegions(
	look: TableLook | undefined,
	rowIndex: number,
	rowCount: number,
	colIndex: number,
	colCount: number,
): string[] {
	const regions: string[] = ['wholeTable'];
	if (!look?.noVBand) regions.push(colIndex % 2 === 0 ? 'band1Vert' : 'band2Vert');
	if (!look?.noHBand) regions.push(rowIndex % 2 === 0 ? 'band1Horz' : 'band2Horz');
	if (look?.lastColumn && colIndex === colCount - 1) regions.push('lastCol');
	if (look?.firstColumn && colIndex === 0) regions.push('firstCol');
	if (look?.lastRow && rowIndex === rowCount - 1) regions.push('lastRow');
	if (look?.firstRow && rowIndex === 0) regions.push('firstRow');
	return regions;
}

/**
 * Resolves the table style's conditional formatting (banding, first/last row/col) for one cell.
 * Direct `tcPr` values on the cell itself always take precedence over this result; callers merge
 * it in as a fallback layer without ever writing it back into the model.
 */
export function resolveTableStyleFormatting(
	tableStyleId: string | undefined,
	catalog: TableStyleCatalog | undefined,
	look: TableLook | undefined,
	rowIndex: number,
	rowCount: number,
	colIndex: number,
	colCount: number,
): TableStyleConditionalFormatting {
	const chain = styleChain(tableStyleId, catalog);
	const regions = applicableRegions(look, rowIndex, rowCount, colIndex, colCount);
	const result: TableStyleConditionalFormatting = {};
	for (const style of chain) {
		if (style.borders) result.borders = { ...result.borders, ...style.borders };
		if (style.shadingFill) result.shadingFill = style.shadingFill;
		for (const region of regions) {
			const formatting = style.conditional[region as keyof typeof style.conditional];
			if (!formatting) continue;
			if (formatting.borders) result.borders = { ...result.borders, ...formatting.borders };
			if (formatting.shadingFill) result.shadingFill = formatting.shadingFill;
			if (formatting.shadingThemeFill) result.shadingThemeFill = formatting.shadingThemeFill;
			if (formatting.run) result.run = { ...result.run, ...formatting.run };
		}
	}
	return result;
}
