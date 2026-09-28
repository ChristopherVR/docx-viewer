// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	AbstractNumDefinition,
	NumberingCatalog,
	NumberingLevelDefinition,
	NumDefinition,
} from './numbering-model.js';
import { expectDefined } from './expect-defined.js';

export type ListKind = 'bullet' | 'decimal';

const BULLET_GLYPHS = ['•', '◦', '▪'];
const DECIMAL_FORMATS = ['decimal', 'lowerLetter', 'lowerRoman'] as const;

function nextId(ids: string[], floor: number): string {
	const max = ids.reduce((max, id) => {
		const value = Number(id);
		return Number.isSafeInteger(value) ? Math.max(max, value) : max;
	}, floor - 1);
	return String(max + 1);
}

function buildLevels(kind: ListKind): Record<number, NumberingLevelDefinition> {
	const levels: Record<number, NumberingLevelDefinition> = {};
	for (let level = 0; level < 9; level++) {
		const base: NumberingLevelDefinition = {
			level,
			start: 1,
			numFmt:
				kind === 'bullet'
					? 'bullet'
					: expectDefined(DECIMAL_FORMATS[level % DECIMAL_FORMATS.length], 'decimal format'),
			lvlText:
				kind === 'bullet'
					? expectDefined(BULLET_GLYPHS[level % BULLET_GLYPHS.length], 'bullet glyph')
					: `%${level + 1}.`,
			indentLeftTwips: 720 * (level + 1),
			hangingTwips: 360,
			suffix: 'tab',
		};
		levels[level] = base;
	}
	return levels;
}

/**
 * Returns a catalog with one additional, independent list definition of the requested kind and
 * its `numId`. Pure: never mutates `catalog`. Existing entries are always reused unchanged, so
 * every call mints a fresh `abstractNum`/`num` pair with its own restart-from-1 counter, matching
 * how Word starts a new list rather than continuing an unrelated one.
 */
export function ensureListDefinition(
	catalog: NumberingCatalog | undefined,
	kind: ListKind,
): { catalog: NumberingCatalog; numId: number } {
	const abstractNums: Record<string, AbstractNumDefinition> = { ...(catalog?.abstractNums ?? {}) };
	const nums: Record<string, NumDefinition> = { ...(catalog?.nums ?? {}) };
	const abstractId = nextId(Object.keys(abstractNums), 0);
	const numId = nextId(Object.keys(nums), 1);
	abstractNums[abstractId] = { id: abstractId, levels: buildLevels(kind) };
	nums[numId] = { id: numId, abstractNumId: abstractId };
	return {
		catalog: { abstractNums, nums, warnings: catalog?.warnings ?? [] },
		numId: Number(numId),
	};
}
