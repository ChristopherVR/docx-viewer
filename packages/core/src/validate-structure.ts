// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Pre-save validation of tables, sections and numbering against the ECMA-376 simple types.
import {
	isStBorder,
	isStJcTable,
	isStMerge,
	isStNumberFormat,
	isStPageOrientation,
	isStSectionMark,
	isStVerticalJc,
} from './generated/wml-simple-types.js';
import type { NumberingCatalog } from './numbering-model.js';
import type { SectionProperties, Table } from './model.js';
import type { TableBorderSide, TableBorders, TableCell } from './table-model.js';
import { validateThemeReference } from './validate-paragraph.js';
import { Checker } from './validate-issues.js';

function validateBorderSide(c: Checker, side: TableBorderSide | undefined): void {
	if (!side) return;
	c.enum('style', side.style, isStBorder, 'ST_Border');
	c.unsigned('sizeEighthPoints', side.sizeEighthPoints, 'ST_EighthPointMeasure');
	c.unsigned('spacePoints', side.spacePoints, 'ST_PointMeasure');
	c.hex('color', side.color);
	c.theme('themeColor', side.themeColor);
}

export function validateBorders(c: Checker, borders: TableBorders | undefined): void {
	if (!borders) return;
	for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const)
		validateBorderSide(c.at(`.${side}`), borders[side]);
}

function validateCell(c: Checker, cell: TableCell): void {
	c.check(
		'gridSpan',
		cell.gridSpan,
		(v) => typeof v === 'number' && Number.isSafeInteger(v) && v >= 1,
		'must be a positive integer (CT_DecimalNumber)',
	);
	c.enum('verticalMerge', cell.verticalMerge, isStMerge, 'ST_Merge');
	c.unsigned('widthTwips', cell.widthTwips, 'ST_TwipsMeasure');
	c.enum('verticalAlign', cell.verticalAlign, isStVerticalJc, 'ST_VerticalJc');
	c.hex('shadingFill', cell.shadingFill);
	validateThemeReference(c, 'shadingThemeFill', cell.shadingThemeFill);
	validateBorders(c.at('.borders'), cell.borders);
	if (cell.margins)
		for (const side of ['top', 'bottom', 'left', 'right'] as const)
			c.at('.margins').signed(side, cell.margins[side], 'ST_SignedTwipsMeasure');
}

/** Validates table-level and cell-level properties; cell paragraphs are walked by the caller. */
export function validateTable(c: Checker, table: Table): void {
	c.unsigned('widthTwips', table.widthTwips, 'ST_TwipsMeasure');
	c.signed('indentTwips', table.indentTwips, 'ST_SignedTwipsMeasure');
	c.check(
		'alignment',
		table.alignment,
		(v) => v === 'left' || v === 'center' || v === 'right',
		'must be "left", "center" or "right"',
	);
	c.enum('justification', table.justification, isStJcTable, 'ST_JcTable');
	table.grid?.forEach((width, index) =>
		c.at(`.grid[${index}]`).unsigned('width', width, 'ST_TwipsMeasure'),
	);
	validateBorders(c.at('.borders'), table.borders);
	if (table.cellMargins)
		for (const side of ['top', 'bottom', 'left', 'right'] as const)
			c.at('.cellMargins').signed(side, table.cellMargins[side], 'ST_SignedTwipsMeasure');
	table.rowProperties?.forEach((row, index) => {
		if (row)
			c.at(`.rowProperties[${index}]`).unsigned('heightTwips', row.heightTwips, 'ST_TwipsMeasure');
	});
	table.rows.forEach((row, r) =>
		row.forEach((cell, k) => validateCell(c.at(`.rows[${r}][${k}]`), cell)),
	);
}

export function validateSection(c: Checker, section: SectionProperties): void {
	c.enum('type', section.type, isStSectionMark, 'ST_SectionMark');
	c.enum('orientation', section.orientation, isStPageOrientation, 'ST_PageOrientation');
	c.enum('verticalAlign', section.verticalAlign, isStVerticalJc, 'ST_VerticalJc');
	for (const key of ['pageWidthTwips', 'pageHeightTwips'] as const)
		c.check(
			key,
			section[key],
			(v) => typeof v === 'number' && Number.isSafeInteger(v) && v > 0,
			'must be a positive integer number of twips (CT_PageSz)',
		);
	for (const key of ['marginTopTwips', 'marginBottomTwips'] as const)
		c.signed(key, section[key], 'ST_SignedTwipsMeasure');
	for (const key of [
		'marginLeftTwips',
		'marginRightTwips',
		'headerDistanceTwips',
		'footerDistanceTwips',
		'gutterTwips',
	] as const)
		c.unsigned(key, section[key], 'ST_TwipsMeasure');
	const columns = c.at('.columns');
	columns.check(
		'count',
		section.columns?.count,
		(v) => typeof v === 'number' && Number.isSafeInteger(v) && v >= 1 && v <= 45,
		'must be an integer from 1 to 45 (CT_Columns/@num)',
	);
	columns.unsigned('spacingTwips', section.columns?.spacingTwips, 'ST_TwipsMeasure');
	section.columns?.widths?.forEach((column, index) => {
		const at = columns.at(`.widths[${index}]`);
		at.unsigned('widthTwips', column.widthTwips, 'ST_TwipsMeasure');
		at.unsigned('spacingTwips', column.spacingTwips, 'ST_TwipsMeasure');
	});
	if (section.lineNumberSettings) {
		const line = c.at('.lineNumberSettings');
		const settings = section.lineNumberSettings;
		line.check(
			'countBy',
			settings.countBy,
			(v) => typeof v === 'number' && Number.isSafeInteger(v) && v >= 1,
			'must be a positive integer (CT_LineNumber/@countBy)',
		);
		line.check(
			'start',
			settings.start,
			(v) => typeof v === 'number' && Number.isSafeInteger(v) && v >= 1,
			'must be a positive visible line number',
		);
		line.check(
			'restart',
			settings.restart,
			(v) => v === 'newPage' || v === 'newSection' || v === 'continuous',
			'must be "newPage", "newSection" or "continuous" (ST_LineNumberRestart)',
		);
		line.unsigned('distanceTwips', settings.distanceTwips, 'ST_TwipsMeasure');
	}
	if (section.pageNumbering) {
		const numbering = c.at('.pageNumbering');
		numbering.enum('format', section.pageNumbering.format, isStNumberFormat, 'ST_NumberFormat');
		numbering.signed('start', section.pageNumbering.start, 'ST_DecimalNumber');
	}
}

export function validateNumberingCatalog(c: Checker, catalog: NumberingCatalog): void {
	const levelOf = (
		at: Checker,
		level: NumberingCatalog['abstractNums'][string]['levels'][number],
	) => {
		at.enum('numFmt', level.numFmt, isStNumberFormat, 'ST_NumberFormat');
		at.signed('start', level.start, 'ST_DecimalNumber');
		at.unsigned('level', level.level, 'ST_DecimalNumber');
		at.check(
			'lvlJc',
			level.lvlJc,
			(v) => v === 'left' || v === 'center' || v === 'right',
			'must be "left", "center" or "right"',
		);
		at.check(
			'suffix',
			level.suffix,
			(v) => v === 'tab' || v === 'space' || v === 'none',
			'must be "tab", "space" or "none"',
		);
		at.check('lvlText', level.lvlText, (v) => typeof v === 'string', 'must be a string');
		at.check(
			'paragraphStyleId',
			level.paragraphStyleId,
			(v) => typeof v === 'string' && v.length > 0,
			'must be a nonempty style ID',
		);
		at.signed('indentLeftTwips', level.indentLeftTwips, 'ST_SignedTwipsMeasure');
		at.unsigned('hangingTwips', level.hangingTwips, 'ST_TwipsMeasure');
		at.unsigned('firstLineTwips', level.firstLineTwips, 'ST_TwipsMeasure');
		at.unsigned('lvlRestart', level.lvlRestart, 'ST_DecimalNumber');
	};
	for (const [id, abstract] of Object.entries(catalog.abstractNums))
		for (const [key, level] of Object.entries(abstract.levels))
			levelOf(new Checker(c.issues, `${c.path}.abstractNums[${id}].levels[${key}]`), level);
	for (const [id, num] of Object.entries(catalog.nums))
		for (const [key, override] of Object.entries(num.levelOverrides ?? {})) {
			const at = new Checker(c.issues, `${c.path}.nums[${id}].levelOverrides[${key}]`);
			at.signed('startOverride', override.startOverride, 'ST_DecimalNumber');
			if (override.lvl) levelOf(at.at('.lvl'), override.lvl);
		}
}
