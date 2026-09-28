// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Shared plumbing for pre-save model validation: issue records, the typed error and the
// primitive schema-value predicates every rule is built from.
import { isStThemeColor } from './generated/wml-simple-types.js';
import { isValidLanguageTag } from './language.js';
import { parseHexColor } from './simple-types.js';

/** One schema-invalid value found in a document model. */
export interface ValidationIssue {
	/** Location of the offending object, e.g. `blocks[2].runs[0]`. */
	path: string;
	/** Model field on that object, e.g. `fontSize`. */
	field: string;
	/** The rejected value. */
	value: unknown;
	/** Human-readable rule that was violated, including the offending value. */
	rule: string;
}

/** Thrown by `saveDocx` (and `assertValidDocumentModel`) when the model cannot be written as valid OOXML. */
export class DocxModelValidationError extends Error {
	readonly issues: readonly ValidationIssue[];
	constructor(issues: readonly ValidationIssue[]) {
		const shown = issues
			.slice(0, 10)
			.map((issue) => `- ${issue.path}.${issue.field}: ${issue.rule}`);
		if (issues.length > shown.length) shown.push(`- ...and ${issues.length - shown.length} more`);
		super(
			`Document model is not valid WordprocessingML (${issues.length} issue${issues.length === 1 ? '' : 's'}); nothing was saved:\n${shown.join('\n')}`,
		);
		this.name = 'DocxModelValidationError';
		this.issues = issues;
	}
}

export const show = (value: unknown): string => {
	if (typeof value === 'string') return JSON.stringify(value);
	if (typeof value === 'number') return String(value);
	try {
		return JSON.stringify(value) ?? String(value);
	} catch {
		return String(value);
	}
};

export const isInteger = (value: unknown): boolean =>
	typeof value === 'number' && Number.isSafeInteger(value);
export const isUnsignedInteger = (value: unknown): boolean =>
	isInteger(value) && (value as number) >= 0;

/** `ST_HexColor` as the model stores it: `auto`, or RRGGBB with an optional leading `#`. */
export function isHexColorInput(value: unknown): boolean {
	if (typeof value !== 'string') return false;
	if (value.startsWith('#')) return /^#[0-9a-fA-F]{6}$/.test(value);
	return parseHexColor(value) !== undefined;
}

const DATE_TIME =
	/^(-?\d{4,})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(\.\d+)?(Z|[+-]\d{2}:\d{2})?$/;
/** `xsd:dateTime` with calendar and clock ranges checked (24:00:00 is the only hour-24 value). */
export function isXsdDateTime(value: unknown): boolean {
	if (typeof value !== 'string') return false;
	const match = DATE_TIME.exec(value);
	if (!match) return false;
	const [year, month, day, hour, minute, second] = match.slice(1, 7).map(Number);
	if (year === 0 || month < 1 || month > 12 || day < 1) return false;
	const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
	const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1];
	if (day > days || minute > 59 || second > 59) return false;
	if (hour === 24) return minute === 0 && second === 0;
	if (hour > 23) return false;
	const zone = /[+-](\d{2}):(\d{2})$/.exec(match[8] ?? '');
	return !zone || (Number(zone[1]) <= 14 && Number(zone[2]) <= 59);
}

/** Collects issues for one object; every check is skipped when the value is `undefined`. */
export class Checker {
	constructor(
		readonly issues: ValidationIssue[],
		readonly path: string,
	) {}
	at(suffix: string): Checker {
		return new Checker(this.issues, `${this.path}${suffix}`);
	}
	fail(field: string, value: unknown, rule: string): void {
		this.issues.push({ path: this.path, field, value, rule });
	}
	check(field: string, value: unknown, test: (v: unknown) => boolean, describe: string): void {
		if (value !== undefined && !test(value))
			this.fail(field, value, `${describe}, got ${show(value)}`);
	}
	enum(field: string, value: unknown, guard: (v: unknown) => boolean, type: string): void {
		this.check(field, value, guard, `must be a valid ${type} value`);
	}
	unsigned(field: string, value: unknown, type: string): void {
		this.check(field, value, isUnsignedInteger, `must be a non-negative integer (${type})`);
	}
	signed(field: string, value: unknown, type: string): void {
		this.check(field, value, isInteger, `must be an integer (${type})`);
	}
	hex(field: string, value: unknown): void {
		this.check(
			field,
			value,
			isHexColorInput,
			'must be ST_HexColor ("auto" or RRGGBB, optional leading #)',
		);
	}
	theme(field: string, value: unknown): void {
		this.enum(field, value, isStThemeColor, 'ST_ThemeColor');
	}
	dateTime(field: string, value: unknown): void {
		this.check(field, value, isXsdDateTime, 'must be an xsd:dateTime (e.g. 2024-01-31T09:30:00Z)');
	}
	language(field: string, value: unknown): void {
		// The empty string clears the attribute on write.
		if (value === undefined || value === '') return;
		if (typeof value !== 'string' || !isValidLanguageTag(value))
			this.fail(field, value, `Invalid BCP 47 language tag: ${String(value)}`);
	}
}
