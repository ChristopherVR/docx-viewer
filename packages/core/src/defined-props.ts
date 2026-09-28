// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
/**
 * Copies own enumerable properties, dropping keys whose value is `undefined`.
 * With `exactOptionalPropertyTypes`, an absent key and an explicit `undefined` are different
 * states, and model equality, JSON round-trips and serializer change detection see the difference.
 */
export function definedProps<T extends object>(
	source: T,
): { [K in keyof T]?: Exclude<T[K], undefined> };
export function definedProps(source: object): object {
	return Object.fromEntries(Object.entries(source).filter(([, value]) => value !== undefined));
}
