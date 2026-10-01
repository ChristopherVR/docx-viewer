import { useData } from 'vitepress';
import type { ComputedRef } from 'vue';
import { computed } from 'vue';

import { en } from './en';
import type { LandingCopy } from './types';

/** Add further locales here (keyed by VitePress `lang`) once translated copy exists. */
const dictionaries: Record<string, LandingCopy> = {
	'en-US': en,
};

/** Returns the landing copy for the active VitePress locale (en fallback). */
export function useLandingCopy(): ComputedRef<LandingCopy> {
	const { lang } = useData();
	return computed(() => dictionaries[lang.value] ?? en);
}

export type { LandingCopy } from './types';
