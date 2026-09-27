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
function cssBorderSide(side: BorderSide | undefined): string | null {
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
	].filter(Boolean);
	return declarations.join(';');
}

export function tableCellStyle(attrs: Record<string, unknown>): string {
	const borders = parseBordersJson(attrs.borders);
	const verticalAlign =
		attrs.verticalAlign === 'top' ? 'top' : attrs.verticalAlign === 'bottom' ? 'bottom' : 'middle';
	const declarations = [
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
