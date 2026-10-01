import { readFile, readdir } from 'node:fs/promises';
import assert from 'node:assert/strict';
const root = new URL('../', import.meta.url);
async function sources(directory) {
	const result = [];
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const file = new URL(entry.name + (entry.isDirectory() ? '/' : ''), directory);
		if (entry.isDirectory()) result.push(...(await sources(file)));
		else if (/\.(ts|svelte)$/.test(entry.name)) result.push(file);
	}
	return result;
}
const legacy = await readFile(new URL('packages/legacy/src/index.ts', root), 'utf8');
assert(legacy.includes('@christophervr/ole2/'), 'Legacy adapter must consume published ole2');
const core = JSON.parse(await readFile(new URL('packages/core/package.json', root), 'utf8'));
assert(!core.dependencies?.['@christophervr/ole2'], 'Modern DOCX core must not depend on ole2');
assert(core.exports['./embedded'], 'Embedded DOCX API must be shared with PowerPoint');
assert(
	core.dependencies?.['@christophervr/ooxml-core'] && Object.keys(core.dependencies).length === 1,
	'docx-core must depend only on @christophervr/ooxml-core',
);
for (const [entry, target] of [
	['index.ts', '@christophervr/ooxml-core/docx'],
	['embedded.ts', '@christophervr/ooxml-core/docx/embedded'],
]) {
	const thin = await readFile(new URL(`packages/core/src/${entry}`, root), 'utf8');
	assert(
		thin.includes(`export * from '${target}';`),
		`docx-core ${entry} must stay a thin re-export of ${target}`,
	);
}
assert(
	(await sources(new URL('packages/core/src/', root))).length === 2,
	'docx-core must hold no logic; it lives in @christophervr/ooxml-core/docx',
);
for (const name of ['core', 'legacy', 'document', 'layout', 'web-component', 'bindings']) {
	for (const file of await sources(new URL(`packages/${name}/src/`, root))) {
		const source = await readFile(file, 'utf8');
		assert(
			!source.includes('@christophervr/ole2/docx'),
			`Modern DOCX must remain in ooxml-core/docx: ${file}`,
		);
		assert(
			!/\/(?:ole2-parser|ole-document-doc)-(?:read|write|cfb|fib|pieces|fkp)[^/]*\.ts$/.test(
				file.pathname,
			),
			`Binary implementation duplicated in viewer: ${file}`,
		);
	}
}
console.log('Verified thin docx-core entry points and published legacy codec boundaries.');
