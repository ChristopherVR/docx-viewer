/** Read `items[index]`, failing the test with a clear message when it is absent. */
export function at<T>(items: readonly T[] | undefined, index: number): T {
	const value = items?.[index];
	if (value === undefined) throw new Error(`Expected an item at index ${index}`);
	return value;
}
