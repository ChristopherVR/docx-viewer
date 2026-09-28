// Read-only visual rendering helpers for table/cell borders, shading and layout.
// Table structure/border/shading editing is not yet supported by the writer (see write-table.ts);
// these attrs are sourced only from the parsed model and are never fed back into edits.
interface BorderSide {
	style?: string;
	sizeEighthPoints?: number;
	color?: string;
}
interface Borders {
	top?: BorderSide;
	bottom?: BorderSide;
	left?: BorderSide;
	right?: BorderSide;
	insideH?: BorderSide;
	insideV?: BorderSide;
}

const NONE_STYLES = new Set(['none', 'nil']);
function cssBorderStyle(style: string | undefined): string {
	if (!style || NONE_STYLES.has(style)) return 'none';
	if (style === 'double') return 'double';
	if (style.toLowerCase().includes('dot')) return 'dotted';
	if (style.toLowerCase().includes('dash')) return 'dashed';
	return 'solid';
}
export function cssBorderSide(side: BorderSide | undefined): string | null {
	if (!side || NONE_STYLES.has(side.style ?? '')) return null;
	const widthPx = side.sizeEighthPoints
		? Math.max(1, Math.round((side.sizeEighthPoints / 8) * (4 / 3)))
		: 1;
	const color = side.color ?? '#000000';
	return `${widthPx}px ${cssBorderStyle(side.style)} ${color}`;
}

export function parseBordersJson(value: unknown): Borders | undefined {
	if (typeof value !== 'string' || !value) return undefined;
	try {
		return JSON.parse(value) as Borders;
	} catch {
		return undefined;
	}
}

const twipsPx = (value: unknown): string | null =>
	Number.isSafeInteger(value) ? `${Number(value) / 15}px` : null;

export function tableStyle(attrs: Record<string, unknown>): string {
	const borders = parseBordersJson(attrs.borders);
	const declarations = [
		'border-collapse:collapse',
		twipsPx(attrs.widthTwips) && `width:${twipsPx(attrs.widthTwips)}`,
		attrs.alignment === 'center' && 'margin-left:auto;margin-right:auto',
		attrs.alignment === 'right' && 'margin-left:auto',
		twipsPx(attrs.indentTwips) &&
			attrs.alignment !== 'center' &&
			`margin-left:${twipsPx(attrs.indentTwips)}`,
		// Cells without their own resolved borders (e.g. a table inserted in the editor) draw the
		// table's inside line; tables with no border information fall back to dashed gridlines.
		borders && `--dve-cell-border:${cssBorderSide(borders.insideH ?? borders.top) ?? 'none'}`,
		// Cells without their own margins use the table's (`w:tblCellMar`).
		typeof attrs.cellMargins === 'string' && `--dve-cell-padding:${marginsCss(attrs.cellMargins)}`,
	].filter(Boolean);
	return declarations.join(';');
}

/** Word's default cell margins: none above/below, 0.075" (108 twips) left and right. */
const DEFAULT_CELL_MARGINS = { top: 0, bottom: 0, left: 108, right: 108 };

function marginsCss(value: unknown): string {
	let margins: Partial<typeof DEFAULT_CELL_MARGINS> = {};
	if (typeof value === 'string')
		try {
			margins = JSON.parse(value) as Partial<typeof DEFAULT_CELL_MARGINS>;
		} catch {
			margins = {};
		}
	const side = (key: keyof typeof DEFAULT_CELL_MARGINS) =>
		`${(margins[key] ?? DEFAULT_CELL_MARGINS[key]) / 15}px`;
	return `${side('top')} ${side('right')} ${side('bottom')} ${side('left')}`;
}

function cellPadding(value: unknown): string {
	return value == null
		? `padding:var(--dve-cell-padding, ${marginsCss(null)})`
		: `padding:${marginsCss(value)}`;
}

export function tableCellStyle(attrs: Record<string, unknown>): string {
	const borders = parseBordersJson(attrs.borders);
	// Word aligns cell content to the top unless `w:vAlign` says otherwise.
	const verticalAlign =
		attrs.verticalAlign === 'center'
			? 'middle'
			: attrs.verticalAlign === 'bottom'
				? 'bottom'
				: 'top';
	const declarations = [
		cellPadding(attrs.margins),
		`vertical-align:${verticalAlign}`,
		twipsPx(attrs.widthTwips) && `width:${twipsPx(attrs.widthTwips)}`,
		typeof attrs.shadingFill === 'string' && `background-color:${attrs.shadingFill}`,
		...(['top', 'bottom', 'left', 'right'] as const).map(
			(side) =>
				borders &&
				(cssBorderSide(borders[side])
					? `border-${side}:${cssBorderSide(borders[side])}`
					: `border-${side}-style:none`),
		),
	].filter(Boolean);
	return declarations.join(';');
}

/**
 * CSS for paragraph borders (`w:pBdr`) and shading: each side's line plus its `w:space` gap as
 * padding, so the text keeps its position and the border sits outside it as in Word.
 */
export function paragraphBoxCss(borders: unknown, shadingFill: unknown): string[] {
	const parsed = (typeof borders === 'string' ? parseBordersJson(borders) : borders) as
		| Record<
				'top' | 'bottom' | 'left' | 'right',
				(BorderSide & { spacePoints?: number }) | undefined
		  >
		| undefined;
	const css: string[] = [];
	for (const side of ['top', 'bottom', 'left', 'right'] as const) {
		const line = cssBorderSide(parsed?.[side]);
		if (!line) continue;
		css.push(
			`border-${side}:${line}`,
			`padding-${side}:${((parsed![side]!.spacePoints ?? 0) * 4) / 3}px`,
		);
	}
	if (typeof shadingFill === 'string' && /^#?[0-9a-f]{6}$/i.test(shadingFill))
		css.push(`background-color:#${shadingFill.replace(/^#/, '')}`);
	return css;
}
