import type { Ligatures } from '@christophervr/docx-core';

/** Explicitly sets every optional class so None cancels inherited font features. */
export function ligatureCss(value: Ligatures): string {
	const enabled = (token: string) => value === 'all' || value.toLowerCase().includes(token);
	return `${enabled('standard') ? 'common-ligatures' : 'no-common-ligatures'} ${enabled('contextual') ? 'contextual' : 'no-contextual'} ${enabled('historical') ? 'historical-ligatures' : 'no-historical-ligatures'} ${enabled('discretional') ? 'discretionary-ligatures' : 'no-discretionary-ligatures'}`;
}
