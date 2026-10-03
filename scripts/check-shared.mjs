import { readFile, readdir } from 'node:fs/promises';
import assert from 'node:assert/strict';
import { workspacePackageNames } from './workspace-packages.mjs';
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
const core = JSON.parse(await readFile(new URL('packages/core/package.json', root), 'utf8'));
assert(!core.dependencies?.['@christophervr/ole2'], 'Modern DOCX core must not depend on ole2');
assert(core.exports['./embedded'], 'Embedded DOCX API must be shared with PowerPoint');
assert(
	core.dependencies?.['ooxml-core'] && Object.keys(core.dependencies).length === 1,
	'docx-core must depend only on ooxml-core',
);
for (const [entry, target] of [
	['index.ts', 'ooxml-core/docx'],
	['embedded.ts', 'ooxml-core/docx/embedded'],
]) {
	const thin = await readFile(new URL(`packages/core/src/${entry}`, root), 'utf8');
	assert(
		thin.includes(`export * from '${target}';`),
		`docx-core ${entry} must stay a thin re-export of ${target}`,
	);
}
assert(
	(await sources(new URL('packages/core/src/', root))).length === 2,
	'docx-core must hold no logic; it lives in ooxml-core/docx',
);
const PUBLISHED = ['core', 'react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];
const INTERNAL = ['web-component', 'bindings'];
for (const name of [...PUBLISHED, ...INTERNAL]) {
	const manifest = JSON.parse(
		await readFile(new URL(`packages/${name}/package.json`, root), 'utf8'),
	);
	assert(
		!JSON.stringify(manifest).includes('@christophervr/ole2'),
		`packages/${name} must not name ole2; the legacy codecs are inlined by ooxml-core`,
	);
	if (INTERNAL.includes(name)) {
		assert.equal(manifest.private, true, `packages/${name} is internal and must be private`);
		continue;
	}
	assert.notEqual(manifest.private, true, `packages/${name} is published and must not be private`);
	for (const dep of Object.keys({ ...manifest.dependencies, ...manifest.peerDependencies })) {
		assert(
			!/^(?:@christophervr\/|docx-(?!core$)|ooxml-(?!core$|ui$))/.test(dep),
			`packages/${name} must not depend on ${dep}; internal packages and ole2 are bundled`,
		);
	}
}
assert.deepEqual(
	await workspacePackageNames(new URL('packages/', root)),
	[...PUBLISHED, ...INTERNAL].sort(),
	'every workspace package must be classified as published or internal',
);
for (const name of [...PUBLISHED, ...INTERNAL]) {
	for (const file of await sources(new URL(`packages/${name}/src/`, root))) {
		const source = await readFile(file, 'utf8');
		assert(
			!source.includes('@christophervr/ole2'),
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
console.log(
	'Verified thin docx-core entry points and that ole2 and legacy codecs stay out of this repository.',
);
