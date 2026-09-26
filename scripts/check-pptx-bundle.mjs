import { readFile, readdir } from 'node:fs/promises';
import assert from 'node:assert/strict';
const root = new URL('../../pptx-viewer-new/packages/core/dist/', import.meta.url);
async function check(directory) {
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const file = new URL(entry.name + (entry.isDirectory() ? '/' : ''), directory);
		if (entry.isDirectory()) {
			await check(file);
			continue;
		}
		if (!/\.(?:m?js|d\.ts)$/.test(entry.name)) continue;
		const source = await readFile(file, 'utf8');
		assert(
			!/(?:from\s*|import\s*\(|require\s*\()\s*['"]@christophervr\/ole2/.test(source),
			`Unbundled local/optional dependency in ${file.pathname}`,
		);
	}
}
await check(root);
const esm = await import(new URL('index.mjs', root).href);
assert(esm.PptxHandler, 'Missing built PowerPoint public API');
console.log(
	'PowerPoint ESM import passes; shared legacy codecs and types are bundled without local runtime imports.',
);
