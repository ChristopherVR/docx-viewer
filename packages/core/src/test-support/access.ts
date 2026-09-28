// Test-only: checked accessors that keep tests readable under noUncheckedIndexedAccess.
// Each helper throws a descriptive error instead of letting `undefined` leak into assertions.
import type { Block, Paragraph, Table } from '../model.js';

/** Returns `items[index]`, failing the test with a clear message when it is missing. */
export function at<T>(items: Readonly<Record<number, T>> | undefined, index: number): T {
	const item = items?.[index];
	if (item === undefined) throw new Error(`Expected an item at index ${index}`);
	return item;
}

/** Returns the only element of a one-element array. */
export function only<T>(items: readonly T[] | undefined): T {
	if (items?.length !== 1) throw new Error(`Expected exactly one item, got ${items?.length}`);
	return at(items, 0);
}

/** Returns the last element of a non-empty array. */
export function lastOf<T>(items: readonly T[] | undefined): T {
	return at(items, (items?.length ?? 0) - 1);
}

/** Narrows a defined value, failing with `what` when it is missing. */
export function must<T>(value: T | null | undefined, what = 'value'): T {
	if (value === undefined || value === null) throw new Error(`Expected ${what} to be defined`);
	return value;
}

/** Narrows a block to a paragraph. */
export function expectParagraph(block: Block | undefined): Paragraph {
	if (block?.type !== 'paragraph') throw new Error('Expected a paragraph block');
	return block;
}

/** Narrows a block to a table. */
export function expectTable(block: Block | undefined): Table {
	if (block?.type !== 'table') throw new Error('Expected a table block');
	return block;
}
