// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	TableConditionalRegion,
	TableStyleCatalog,
	TableStyleConditionalFormatting,
	TableStyleDefinition,
} from './table-model.js';
import {
	children,
	first,
	getW,
	named,
	parseXml,
	type XmlDocument,
	type XmlElement,
} from './xml.js';
import { parseRunProperties } from './run-properties.js';
import { parseOnOff } from './simple-types.js';
import { parseShadingFill, parseShadingThemeFill, parseTableBorders } from './table-borders.js';

const REGIONS: TableConditionalRegion[] = [
	'wholeTable',
	'firstRow',
	'lastRow',
	'firstCol',
	'lastCol',
	'band1Horz',
	'band2Horz',
	'band1Vert',
	'band2Vert',
];

function tableStyleElements(document: XmlDocument): XmlElement[] {
	return Array.from(document.getElementsByTagName('*')).filter(
		(element): element is XmlElement =>
			named(element, 'style') && getW(element, 'type') === 'table',
	);
}

function conditionalFormatting(block: XmlElement): TableStyleConditionalFormatting {
	const result: TableStyleConditionalFormatting = {};
	const tcPr = first(block, 'tcPr');
	const borders = parseTableBorders(first(tcPr, 'tcBorders'));
	if (borders) result.borders = borders;
	const shd = first(tcPr, 'shd');
	const fill = parseShadingFill(shd);
	if (fill) result.shadingFill = fill;
	const themeFill = parseShadingThemeFill(shd);
	if (themeFill) result.shadingThemeFill = themeFill;
	const run = parseRunProperties(first(block, 'rPr'));
	if (Object.keys(run).length) result.run = run;
	return result;
}

/** Parses `styles.xml` table styles, including `w:tblStylePr` conditional formatting regions. */
export function parseTableStyleCatalog(xml: string): TableStyleCatalog {
	const document = parseXml(xml);
	const styles = Object.create(null) as Record<string, TableStyleDefinition>;
	for (const element of tableStyleElements(document)) {
		const id = getW(element, 'styleId');
		if (!id) continue;
		const basedOn = getW(first(element, 'basedOn'), 'val');
		const isDefault = parseOnOff(getW(element, 'default'));
		const name = getW(first(element, 'name'), 'val');
		const tblPr = first(element, 'tblPr');
		const borders = parseTableBorders(first(tblPr, 'tblBorders'));
		const shadingFill = parseShadingFill(first(tblPr, 'shd'));
		const conditional: Partial<Record<TableConditionalRegion, TableStyleConditionalFormatting>> =
			{};
		for (const block of children(element, 'tblStylePr')) {
			const type = getW(block, 'type');
			if (type && (REGIONS as string[]).includes(type))
				conditional[type as TableConditionalRegion] = conditionalFormatting(block);
		}
		styles[id] = {
			id,
			...(name ? { name } : {}),
			...(basedOn ? { basedOn } : {}),
			...(isDefault === undefined ? {} : { isDefault }),
			...(borders ? { borders } : {}),
			...(shadingFill ? { shadingFill } : {}),
			conditional,
		};
	}
	const warnings: string[] = [];
	for (const style of Object.values(styles)) {
		const seen = new Set<string>([style.id]);
		let parent = style.basedOn;
		while (parent && styles[parent]) {
			if (seen.has(parent)) {
				warnings.push(
					`Table style inheritance cycle detected at "${parent}"; cyclic inheritance is ignored.`,
				);
				break;
			}
			seen.add(parent);
			parent = styles[parent]?.basedOn;
		}
	}
	return { styles, warnings: [...new Set(warnings)] };
}
