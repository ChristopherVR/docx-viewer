import type { Block, Paragraph, Table } from '@christophervr/docx-core';
import { expectDefined } from './defined';

/** The element at `index`, failing the test with a clear message when the list is too short. */
export function at<T>(list: ArrayLike<T>, index: number): T {
	return expectDefined(list[index], `item ${index} of ${list.length}`);
}

/** Asserts a value is present (not undefined/null) and returns it with a narrowed type. */
export function must<T>(value: T | null | undefined, what = 'value'): T {
	return expectDefined(value, what);
}

/** The paragraph block at `index`, failing the test when it is missing or is a table. */
export function paragraphAt(blocks: readonly Block[], index: number): Paragraph {
	const block = at(blocks, index);
	if (block.type !== 'paragraph') throw new Error(`Expected block ${index} to be a paragraph.`);
	return block;
}

/** The table block at `index`, failing the test when it is missing or is a paragraph. */
export function tableAt(blocks: readonly Block[], index: number): Table {
	const block = at(blocks, index);
	if (block.type !== 'table') throw new Error(`Expected block ${index} to be a table.`);
	return block;
}
