import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { LegacyDocError, loadLegacyDoc } from '../index';
import { buildOle2 } from '@christophervr/ole2/ole2-parser-write';
import { parseOle2 } from '@christophervr/ole2/ole2-parser-read';
import { ENTRY_TYPE_ROOT, ENTRY_TYPE_STREAM } from '@christophervr/ole2/ole2-parser-types';

const fixture = new URL('./fixtures/ole-word-97.doc', import.meta.url);

describe('legacy .doc', () => {
	it('extracts Word 97 text and preserves original bytes on a no-op save', async () => {
		const bytes = new Uint8Array(await readFile(fixture));
		const loaded = await loadLegacyDoc(bytes);
		expect(loaded.model.blocks.length).toBeGreaterThan(0);
		expect(loaded.model.blocks[0]?.type).toBe('paragraph');
		expect(loaded.model.warnings[0]).toMatch(/main-body text only/);
		expect(await loaded.save()).toEqual(bytes);
	});

	it('rejects truncated and non-OLE inputs clearly', async () => {
		await expect(loadLegacyDoc(new Uint8Array([0xd0, 0xcf]))).rejects.toThrow(
			/Invalid or truncated/,
		);
		await expect(loadLegacyDoc(new Uint8Array(512))).rejects.toThrow(LegacyDocError);
	});

	it('rejects structural edits beyond the supported subset', async () => {
		const loaded = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
		loaded.model.blocks.push({ type: 'paragraph', id: 'extra', runs: [{ text: 'new' }] });
		await expect(loaded.save()).rejects.toThrow(/Adding or removing paragraphs/);
	});

	it('rejects formatting and page-layout edits it cannot represent safely', async () => {
		const formatted = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
		const first = formatted.model.blocks[0];
		if (first?.type !== 'paragraph') throw new Error('fixture has no paragraph');
		first.runs[0]!.bold = true;
		await expect(formatted.save()).rejects.toThrow(/formatting edits are unsupported/);

		const pageChanged = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
		pageChanged.model.page.marginTop += 1;
		await expect(pageChanged.save()).rejects.toThrow(/page layout/);
	});

	it('reports encrypted legacy Word documents explicitly', async () => {
		const source = new Uint8Array(await readFile(fixture));
		const parsed = parseOle2(
			source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) as ArrayBuffer,
		);
		const streams = new Map<string, Uint8Array>();
		for (const entry of parsed.entries) {
			if (entry.type === ENTRY_TYPE_STREAM) {
				const stream = parsed.getStream(entry.name);
				if (stream) streams.set(entry.name, stream);
			}
		}
		const word = streams.get('WordDocument');
		if (!word) throw new Error('fixture has no WordDocument stream');
		const flags = new DataView(word.buffer, word.byteOffset, word.byteLength).getUint16(0x0a, true);
		new DataView(word.buffer, word.byteOffset, word.byteLength).setUint16(
			0x0a,
			flags | (1 << 8),
			true,
		);
		streams.set('WordDocument', word);
		const clsid = parsed.entries.find((entry) => entry.type === ENTRY_TYPE_ROOT)?.clsid;
		const encrypted = new Uint8Array(buildOle2(streams, clsid));
		await expect(loadLegacyDoc(encrypted)).rejects.toThrow(/Encrypted legacy Word/);
	});

	it.each([
		['language', 'ar-SA'],
		['eastAsiaLanguage', 'ja-JP'],
		['bidiLanguage', 'he-IL'],
		['rtl', false],
		['strike', true],
		['highlight', 'yellow'],
		['verticalAlign', 'superscript'],
	])('rejects unsupported run property %s instead of silently dropping it', async (key, value) => {
		const loaded = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		Object.assign(paragraph.runs[0], { [String(key)]: value });
		await expect(loaded.save()).rejects.toThrow(/formatting edits are unsupported/);
	});

	it('rejects paragraph direction, spacing and new line breaks in legacy DOC output', async () => {
		for (const change of [{ direction: 'rtl' }, { lineSpacingTwips: 480 }]) {
			const loaded = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
			Object.assign(loaded.model.blocks[0], change);
			await expect(loaded.save()).rejects.toThrow(/formatting edits are unsupported/);
		}
		const loaded = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		paragraph.runs[0].text += '\nAnother line';
		await expect(loaded.save()).rejects.toThrow(/line breaks/);
	});

	it('saves supported paragraph text edits and can reopen them', async () => {
		const loaded = await loadLegacyDoc(new Uint8Array(await readFile(fixture)));
		const first = loaded.model.blocks[0];
		if (first?.type !== 'paragraph') throw new Error('fixture has no paragraph');
		first.runs[0]!.text += ' edited';
		const saved = await loaded.save();
		const reopened = await loadLegacyDoc(saved);
		const para = reopened.model.blocks[0];
		expect(para?.type === 'paragraph' ? para.runs.map((run) => run.text).join('') : '').toContain(
			'edited',
		);
	});
});
