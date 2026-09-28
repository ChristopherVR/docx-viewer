// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph, ParagraphStyleCatalog } from './model.js';
import type {
	NumberingCatalog,
	NumberingLevelDefinition,
	ParagraphListLabel,
} from './numbering-model.js';
import { definedProps } from './defined-props.js';
import { resolveNumberingLevel } from './numbering-parse.js';
import { resolveStyleNumbering } from './paragraph-styles.js';
import {
	cardinalWords,
	letterLabel,
	ordinalNumeral,
	ordinalWords,
	romanNumeral,
} from './numbering-text.js';

/** Formats one placeholder value using a level's numbering format. */
export function formatListNumber(numFmt: string, value: number): string {
	switch (numFmt) {
		case 'decimal':
			return String(value);
		case 'decimalZero':
			return value < 10 && value >= 0 ? `0${value}` : String(value);
		case 'upperRoman':
			return romanNumeral(value);
		case 'lowerRoman':
			return romanNumeral(value).toLowerCase();
		case 'upperLetter':
			return letterLabel(value).toUpperCase();
		case 'lowerLetter':
			return letterLabel(value);
		case 'ordinal':
			return ordinalNumeral(value);
		case 'ordinalText':
			return ordinalWords(value);
		case 'cardinalText':
			return cardinalWords(value);
		case 'none':
			return '';
		case 'bullet':
			return '';
		default:
			// Unsupported/legacy formats (e.g. chicago, hex, chineseCounting) fall back to Decimal Number.
			return String(value);
	}
}

/** Resolves the paragraph's effective `{ numId, level }`, direct or inherited through its style. */
export function resolveParagraphNumbering(
	paragraph: Paragraph,
	styleCatalog: ParagraphStyleCatalog | undefined,
): { numId: number; level: number } | undefined {
	if (paragraph.numbering)
		// `numId="0"` is Word's explicit "no numbering" override; it blocks style inheritance
		// without itself producing a marker.
		return paragraph.numbering.numId > 0 ? paragraph.numbering : undefined;
	if (!styleCatalog) return undefined;
	return resolveStyleNumbering(paragraph.style, styleCatalog);
}

function flattenParagraphs(blocks: Block[]): Paragraph[] {
	const result: Paragraph[] = [];
	for (const block of blocks) {
		if (block.type === 'paragraph') result.push(block);
		else for (const row of block.rows) for (const cell of row) result.push(...cell.paragraphs);
	}
	return result;
}

class NumberingCounters {
	private counts = new Map<string, Map<number, number>>();
	private started = new Map<string, Set<number>>();

	private forNum(numId: string): Map<number, number> {
		let counts = this.counts.get(numId);
		if (!counts) {
			counts = new Map();
			this.counts.set(numId, counts);
		}
		return counts;
	}

	private startedLevels(numId: string): Set<number> {
		let started = this.started.get(numId);
		if (!started) {
			started = new Set();
			this.started.set(numId, started);
		}
		return started;
	}

	/**
	 * Advances the counter for `numId`/`level` and restarts every deeper level, matching Word's
	 * default multilevel behavior. `lvlRestart` is parsed and preserved on the model for
	 * round-tripping, but this simplified counter always restarts on any shallower-level change
	 * rather than modeling `lvlRestart`'s non-default cascade.
	 */
	advance(numId: string, level: number, def: NumberingLevelDefinition): number {
		const counts = this.forNum(numId);
		const started = this.startedLevels(numId);
		const next = started.has(level) ? (counts.get(level) ?? def.start) + 1 : def.start;
		counts.set(level, next);
		started.add(level);
		for (const deeper of [...started]) {
			if (deeper <= level) continue;
			counts.delete(deeper);
			started.delete(deeper);
		}
		return next;
	}

	/** Current value for an ancestor placeholder, defaulting to its own start when never instantiated. */
	current(numId: string, level: number, start: number): number {
		return this.forNum(numId).get(level) ?? start;
	}
}

function substituteLvlText(
	lvlText: string,
	numId: string,
	level: number,
	catalog: NumberingCatalog,
	counters: NumberingCounters,
	forceDecimal: boolean,
): string {
	return lvlText.replace(/%([1-9])/g, (match, digit: string) => {
		const ancestorLevel = Number(digit) - 1;
		if (ancestorLevel > level) return match;
		const ancestorDef = resolveNumberingLevel(catalog, numId, ancestorLevel);
		const start = ancestorDef?.start ?? 1;
		const value = counters.current(numId, ancestorLevel, start);
		const fmt = forceDecimal ? 'decimal' : (ancestorDef?.numFmt ?? 'decimal');
		return formatListNumber(fmt, value);
	});
}

/**
 * Computes list markers for every numbered/bulleted paragraph in document order, tracking
 * counter state across the document per `numId`/level. Purely derived from the model; never
 * mutates it and is never itself part of the saved document.
 */
export function computeListLabels(model: DocumentModel): Map<string, ParagraphListLabel> {
	const catalog = model.numberingCatalog;
	const labels = new Map<string, ParagraphListLabel>();
	if (!catalog) return labels;
	const counters = new NumberingCounters();
	for (const paragraph of flattenParagraphs(model.blocks)) {
		const numbering = resolveParagraphNumbering(paragraph, model.paragraphStyles);
		if (!numbering) continue;
		const numId = String(numbering.numId);
		const level = Math.min(8, Math.max(0, numbering.level));
		const def = resolveNumberingLevel(catalog, numId, level);
		if (!def) continue;
		if (def.numFmt === 'none') continue;
		counters.advance(numId, level, def);
		const text =
			def.numFmt === 'bullet'
				? def.lvlText
				: substituteLvlText(def.lvlText, numId, level, catalog, counters, Boolean(def.isLgl));
		labels.set(paragraph.id, {
			numId,
			level,
			text,
			suffix: def.suffix ?? 'tab',
			...definedProps({
				indentLeftTwips: def.indentLeftTwips,
				hangingTwips: def.hangingTwips,
				firstLineTwips: def.firstLineTwips,
			}),
		});
	}
	return labels;
}

/**
 * Bullets in Word's Symbol and Wingdings fonts use private-use code points (U+F0xx); without those
 * fonts they render as missing glyphs, so labels are shown with their Unicode equivalents.
 */
const SYMBOL_BULLETS: Record<string, string> = {
	'': '•',
	'': '▪',
	'': '➢',
	'': '❖',
	'': '✓',
	'': '■',
	'': '❑',
	'': '➩',
	'': '○',
};
export function displayListLabel(text: string): string {
	return text.replace(/[-]/g, (char) => SYMBOL_BULLETS[char] ?? '•');
}
