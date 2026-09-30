import { isLigatures } from '@christophervr/docx-core';
import { ligatureCss } from '@christophervr/docx-layout';
export { ligatureCss } from '@christophervr/docx-layout';

/** Explicitly sets every optional class so a direct None cancels inherited font features. */
export const ligatureStyle = (value: unknown): string =>
	isLigatures(value) ? `font-variant-ligatures:${ligatureCss(value)}` : '';
