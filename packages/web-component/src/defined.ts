/** Narrows a value the caller knows is present (an internal invariant), failing loudly with context if not. */
export function expectDefined<T>(value: T | null | undefined, context: string): T {
	if (value === undefined || value === null) throw new Error(`Expected ${context} to be present.`);
	return value;
}
