import type { DocumentModel } from '@christophervr/docx-core';
import { adaptDocumentModel } from './adapter.js';
import { layoutSections } from './page-flow.js';
import { FLOAT_WRAP_NOTE, positionFloats } from './floats.js';
import type { LayoutDocumentInput } from './input.js';
import type { LayoutResult } from './result.js';
import type { TextMeasurer } from './measure.js';

/** Paginates an engine-native `LayoutDocumentInput` (bypasses the DocumentModel adapter). */
export function layoutDocument(input: LayoutDocumentInput, measurer: TextMeasurer): LayoutResult {
	const result = layoutSections(input, measurer);
	if (!positionFloats(input, result.pages)) return result;
	return { ...result, approximations: [...new Set([...result.approximations, FLOAT_WRAP_NOTE])] };
}

/** Paginates a `docx-core` `DocumentModel` directly: `adaptDocumentModel` + `layoutDocument`. */
export function layoutDocumentModel(model: DocumentModel, measurer: TextMeasurer): LayoutResult {
	const approximations: string[] = [];
	const input = adaptDocumentModel(model, (message) => approximations.push(message));
	const result = layoutDocument(input, measurer);
	return {
		pages: result.pages,
		approximations: [...new Set([...approximations, ...result.approximations])],
	};
}
