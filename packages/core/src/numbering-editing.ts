// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type {
	AbstractNumDefinition,
	NumberingCatalog,
	NumberingLevelDefinition,
	NumDefinition,
} from './numbering-model.js';
import { expectDefined } from './expect-defined.js';
import { signedTwips, twips } from './units.js';
import { Checker, DocxModelValidationError, type ValidationIssue } from './validate-issues.js';
import { validateNumberingCatalog } from './validate-structure.js';

/**
 * `multilevel` numbers 1., 1.1., 1.1.1. (every level shows its ancestors); `outline` cycles
 * 1., a), i. by level, like Word's plain outline lists.
 */
export type ListKind = 'bullet' | 'decimal' | 'multilevel' | 'outline';
const OUTLINE_FORMATS = [
	['decimal', '.'],
	['lowerLetter', ')'],
	['lowerRoman', '.'],
] as const;

function levelText(kind: ListKind, level: number): string {
	if (kind === 'bullet')
		return expectDefined(BULLET_GLYPHS[level % BULLET_GLYPHS.length], 'bullet glyph');
	if (kind === 'multilevel')
		return `${Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join('.')}.`;
	const [, close] = expectDefined(OUTLINE_FORMATS[level % 3], 'outline format');
	return `%${level + 1}${close}`;
}
function levelFormat(kind: ListKind, level: number) {
	if (kind === 'bullet') return 'bullet' as const;
	if (kind === 'decimal')
		return expectDefined(DECIMAL_FORMATS[level % DECIMAL_FORMATS.length], 'decimal format');
	return kind === 'multilevel'
		? ('decimal' as const)
		: expectDefined(OUTLINE_FORMATS[level % 3], 'outline format')[0];
}

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
			numFmt: levelFormat(kind, level),
			lvlText: kind === 'decimal' ? `%${level + 1}.` : levelText(kind, level),
			indentLeftTwips: signedTwips(720 * (level + 1)),
			hangingTwips: twips(kind === 'multilevel' ? 720 : 360),
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
	return createListDefinition(catalog, Object.values(buildLevels(kind)));
}

/** Adds a validated independent definition without modifying any imported list. */
export function createListDefinition(
	catalog: NumberingCatalog | undefined,
	levels: readonly NumberingLevelDefinition[],
): { catalog: NumberingCatalog; numId: number } {
	if (
		!levels.length ||
		levels.length > 9 ||
		new Set(levels.map((level) => level.level)).size !== levels.length ||
		levels.some((level) => !Number.isInteger(level.level) || level.level < 0 || level.level > 8)
	)
		throw new Error('A list must have 1 to 9 distinct levels numbered 0 through 8.');
	const copied = Object.fromEntries(levels.map((level) => [level.level, { ...level }]));
	const issues: ValidationIssue[] = [];
	validateNumberingCatalog(new Checker(issues, 'numberingCatalog'), {
		abstractNums: { '0': { id: '0', levels: copied } },
		nums: {},
		warnings: [],
	});
	if (issues.length) throw new DocxModelValidationError(issues);
	const abstractNums: Record<string, AbstractNumDefinition> = { ...(catalog?.abstractNums ?? {}) };
	const nums: Record<string, NumDefinition> = { ...(catalog?.nums ?? {}) };
	const abstractId = nextId(Object.keys(abstractNums), 0);
	const numId = nextId(Object.keys(nums), 1);
	abstractNums[abstractId] = { id: abstractId, levels: copied };
	nums[numId] = { id: numId, abstractNumId: abstractId };
	return {
		catalog: { abstractNums, nums, warnings: catalog?.warnings ?? [] },
		numId: Number(numId),
	};
}
