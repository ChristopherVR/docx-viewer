import {
	createDocument,
	type DocumentModel,
	type LoadedDocument,
	type Paragraph,
} from '@christophervr/docx-core';
import {
	readOleDocParagraphs,
	writeOleDocParagraphEdit,
} from '@christophervr/ole2/ole-document-doc-editor';
import { unwrapDocBytes } from '@christophervr/ole2/ole-document-doc-cfb';
import { readDocFib } from '@christophervr/ole2/ole-document-doc-fib';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';

export class LegacyDocError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'LegacyDocError';
	}
}

function toBytes(input: Uint8Array | ArrayBuffer): Uint8Array {
	return input instanceof Uint8Array ? input.slice() : new Uint8Array(input.slice(0));
}

function paragraphText(model: DocumentModel): string[] {
	if (model.paragraphStyles)
		throw new LegacyDocError('Paragraph style catalogs are unsupported in legacy .doc save.');
	if (model.blocks.some((block) => block.type !== 'paragraph')) {
		throw new LegacyDocError(
			'Legacy .doc save supports paragraph text only; tables are unsupported.',
		);
	}
	return (model.blocks as Paragraph[]).map((p) => {
		if (
			Object.entries(p).some(
				([key, value]) => !['type', 'id', 'runs'].includes(key) && value !== undefined,
			) ||
			p.runs.some((run) =>
				Object.entries(run).some(([key, value]) => key !== 'text' && value !== undefined),
			)
		) {
			throw new LegacyDocError(
				'Legacy .doc save supports plain paragraph text only; formatting edits are unsupported.',
			);
		}
		return p.runs.map((run) => run.text).join('');
	});
}

function makeModel(paragraphs: string[]): DocumentModel {
	// A .doc's styles are not modeled, so the new-document default catalogs do not apply.
	const {
		paragraphStyles: _paragraphStyles,
		characterStyles: _characterStyles,
		tableStyles: _tableStyles,
		...model
	} = createDocument();
	model.blocks = paragraphs.map((text, index) => ({
		type: 'paragraph',
		id: `p${index + 1}`,
		runs: [{ text }],
	}));
	if (model.blocks.length === 0)
		model.blocks.push({ type: 'paragraph', id: 'p1', runs: [{ text: '' }] });
	model.warnings.push(
		'Legacy .doc import extracts main-body text only. Layout, tables, headers, footers, images, fields, and most formatting are not represented.',
	);
	return model;
}

/** Load Word 97-2003 binary .doc. This API preserves source bytes for no-op saves. */
export async function loadLegacyDoc(input: Uint8Array | ArrayBuffer): Promise<LoadedDocument> {
	const original = toBytes(input);
	let streams: string[] = [];
	try {
		const parsed = parseOle2(
			original.buffer.slice(
				original.byteOffset,
				original.byteOffset + original.byteLength,
			) as ArrayBuffer,
		);
		streams = parsed.entries.map((entry) => entry.name);
	} catch (error) {
		throw new LegacyDocError(
			`Invalid or truncated OLE compound file: ${error instanceof Error ? error.message : 'parse failed'}`,
		);
	}
	if (streams.includes('EncryptionInfo') || streams.includes('EncryptedPackage')) {
		throw new LegacyDocError('Encrypted Office files are not supported.');
	}
	const cfb = unwrapDocBytes(original);
	if (cfb && (readDocFib(cfb.wordDocBytes).flags1 & (1 << 8)) !== 0) {
		throw new LegacyDocError('Encrypted legacy Word documents are not supported.');
	}
	const paragraphs = readOleDocParagraphs(original);
	if (!paragraphs) {
		throw new LegacyDocError(
			streams.includes('WordDocument')
				? 'Unsupported or malformed legacy Word document.'
				: 'OLE file does not contain a legacy Word document.',
		);
	}
	const model = makeModel(paragraphs);
	const initialPage = JSON.stringify(model.page);
	const initial = JSON.stringify(paragraphText(model));
	return {
		model,
		async save(nextModel = model) {
			if (JSON.stringify(nextModel.page) !== initialPage) {
				throw new LegacyDocError('Changing page layout in legacy .doc files is unsupported.');
			}
			const updated = paragraphText(nextModel);
			if (JSON.stringify(updated) === initial) return original.slice();
			if (updated.length !== paragraphs.length)
				throw new LegacyDocError(
					'Adding or removing paragraphs in legacy .doc files is unsupported.',
				);
			let output = original;
			for (let i = 0; i < updated.length; i++) {
				if (updated[i] === paragraphs[i]) continue;
				if (/[\r\n]/u.test(updated[i]!))
					throw new LegacyDocError(
						'Adding line breaks in legacy .doc paragraph text is unsupported.',
					);
				const result = writeOleDocParagraphEdit(output, i, updated[i]!);
				if (result === output)
					throw new LegacyDocError(
						'This .doc uses structures that prevent safe paragraph text editing.',
					);
				output = result;
			}
			return output;
		},
	};
}
