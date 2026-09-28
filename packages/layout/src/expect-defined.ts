/** Assert an indexed read is present where an invariant guarantees it. */
export function expectDefined<T>(value: T | undefined, context: string): T {
	if (value === undefined) throw new Error(`Layout invariant violated: ${context}`);
	return value;
}
