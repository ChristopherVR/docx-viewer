// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Collects parse warnings raised deep inside pure attribute parsers (invalid enumeration values)
// without threading a context through every call. The sink is scoped to synchronous sections so
// concurrent asynchronous loads never mix warnings.
let sink: string[] | undefined;

/** Runs `fn` with `warnings` as the active sink; nesting restores the previous sink. */
export function withParseWarnings<T>(warnings: string[] | undefined, fn: () => T): T {
	const previous = sink;
	sink = warnings;
	try {
		return fn();
	} finally {
		sink = previous;
	}
}

/** Records a de-duplicated warning when a sink is active. */
export function reportParseWarning(message: string): void {
	if (sink && !sink.includes(message)) sink.push(message);
}

/**
 * Validates a raw attribute against an XSD enumeration guard. Missing values return undefined
 * silently; invalid ones are dropped with a warning naming the attribute.
 */
export function enumValue<T extends string>(
	guard: (value: unknown) => value is T,
	raw: string | null | undefined,
	label: string,
): T | undefined {
	if (raw === null || raw === undefined || raw === '') return undefined;
	if (guard(raw)) return raw;
	reportParseWarning(`Ignored invalid ${label} value “${raw}”; it is not allowed by the schema.`);
	return undefined;
}
