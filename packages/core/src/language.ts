/** Accepts structurally safe BCP 47 language tags for Word language metadata. */
export function isValidLanguageTag(value: string): boolean {
	if (value.length > 255) return false;
	if (/^x(?:-[A-Za-z0-9]{1,8})+$/i.test(value)) return true;
	return /^(?:[A-Za-z]{2,8})(?:-[A-Za-z0-9]{1,8})*$/.test(value);
}
