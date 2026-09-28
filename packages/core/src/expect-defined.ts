// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Typed presence assertion for indexed reads whose presence is guaranteed by an invariant.

/** Returns `value` when defined, otherwise throws a descriptive error naming the broken invariant. */
export function expectDefined<T>(value: T | undefined | null, context: string): T {
	if (value === undefined || value === null) {
		throw new Error(`Internal invariant violated: ${context} is missing.`);
	}
	return value;
}
