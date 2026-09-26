// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { loadDocx, type Paragraph } from './index.js';

/** Embedded-document API shared with PowerPoint. Only direct body paragraphs are exposed. */
export async function readOleDocumentParagraphs(bytes: Uint8Array): Promise<string[] | undefined> {
	try {
		const { model } = await loadDocx(bytes);
		return model.blocks
			.filter((block): block is Paragraph => block.type === 'paragraph')
			.map((paragraph) => paragraph.runs.map((run) => run.text).join(''));
	} catch {
		return undefined;
	}
}

/** Preserve PowerPoint's fail-closed contract and first-run formatting behavior. */
export async function writeOleDocumentParagraphEdit(
	bytes: Uint8Array,
	paragraphIndex: number,
	text: string,
): Promise<Uint8Array> {
	try {
		if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0) return bytes;
		const session = await loadDocx(bytes);
		const paragraph = session.model.blocks.filter(
			(block): block is Paragraph => block.type === 'paragraph',
		)[paragraphIndex];
		if (!paragraph) return bytes;
		if (paragraph.runs.map((run) => run.text).join('') === text) return bytes;
		paragraph.runs = [{ ...paragraph.runs[0], text }];
		return await session.save();
	} catch {
		return bytes;
	}
}
