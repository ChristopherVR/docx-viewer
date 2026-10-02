import { isLigatures } from 'docx-core';
import { ligatureCss } from '@christophervr/ooxml-core/docx/layout';
export { ligatureCss } from '@christophervr/ooxml-core/docx/layout';

/** Explicitly sets every optional class so a direct None cancels inherited font features. */
export const ligatureStyle = (value: unknown): string =>
	isLigatures(value) ? `font-variant-ligatures:${ligatureCss(value)}` : '';
