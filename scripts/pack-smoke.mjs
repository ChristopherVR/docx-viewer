import assert from 'node:assert/strict';
import { cp, mkdtemp, mkdir, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { compile } from 'svelte/compiler';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const work = await mkdtemp(path.join(tmpdir(), 'docx-package-smoke-'));
const npmCli =
	process.platform === 'win32'
		? path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
		: 'npm';
const run = (command, args, options = {}) => {
	const executable = command === 'npm' && process.platform === 'win32' ? process.execPath : command;
	const executableArgs =
		command === 'npm' && process.platform === 'win32' ? [npmCli, ...args] : args;
	const result = spawnSync(executable, executableArgs, { cwd: root, encoding: 'utf8', ...options });
	if (result.status !== 0)
		throw new Error(
			`${command} ${args.join(' ')} failed${result.error ? `: ${result.error.message}` : ''}\n${result.stdout}\n${result.stderr}`,
		);
	return result.stdout;
};
async function assertPublishedImports(directory, packageName) {
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const target = path.join(directory, entry.name);
		if (entry.isDirectory()) await assertPublishedImports(target, packageName);
		else if (/\.(?:js|d\.ts)$/.test(entry.name)) {
			const contents = await readFile(target, 'utf8');
			assert(
				!/@docx-viewer\/|@office-viewers\/ole2/.test(contents),
				`${packageName} contains an unpublished internal package import in ${entry.name}`,
			);
		}
	}
}

try {
	const modules = path.join(work, 'node_modules');
	const scope = path.join(modules, '@christophervr');
	const tarballs = [];
	const peers = new Map();
	await mkdir(scope, { recursive: true });
	// `@christophervr/ooxml-core` is not published yet: pack the sibling checkout and make every
	// package that depends on it by `file:` path depend on that tarball instead.
	const corePacked = JSON.parse(
		run('npm', ['pack', '--json', '--pack-destination', work, path.resolve(root, '../ooxml-core')]),
	)[0];
	const coreTarball = path.join(work, corePacked.filename);
	tarballs.push(coreTarball);
	const coreTarget = path.join(scope, 'ooxml-core');
	await mkdir(coreTarget, { recursive: true });
	run('tar', ['-xzf', coreTarball, '--strip-components=1', '-C', coreTarget]);
	for (const name of [
		'core',
		'legacy',
		'document',
		'layout',
		'web-component',
		'bindings',
		'viewer',
	]) {
		let packageDir = path.join(root, 'packages', name);
		const sourceManifest = JSON.parse(
			await readFile(path.join(packageDir, 'package.json'), 'utf8'),
		);
		if (
			String(sourceManifest.dependencies?.['@christophervr/ooxml-core'] ?? '').startsWith('file:')
		) {
			const copy = path.join(work, `pack-${name}`);
			await mkdir(copy, { recursive: true });
			for (const entry of ['dist', 'README.md', 'LICENSE', 'NOTICE'])
				await cp(path.join(packageDir, entry), path.join(copy, entry), {
					recursive: true,
				}).catch(() => {});
			sourceManifest.dependencies['@christophervr/ooxml-core'] = `file:${coreTarball}`;
			await writeFile(path.join(copy, 'package.json'), JSON.stringify(sourceManifest, null, '	'));
			packageDir = copy;
		}
		const packed = JSON.parse(
			run('npm', ['pack', '--json', '--pack-destination', work, packageDir]),
		)[0];
		const file = path.join(work, packed.filename);
		tarballs.push(file);
		const target = path.join(
			scope,
			name === 'web-component' ? 'docx-web-component' : `docx-${name}`,
		);
		await mkdir(target, { recursive: true });
		run('tar', ['-xzf', file, '--strip-components=1', '-C', target]);
		assert(
			packed.files.some((entry) => entry.path === 'dist/index.js'),
			`${packed.name} has no built JavaScript entry`,
		);
		assert(
			packed.files.some((entry) => entry.path === 'dist/index.d.ts'),
			`${packed.name} has no TypeScript declarations`,
		);
		const manifest = JSON.parse(await readFile(path.join(target, 'package.json'), 'utf8'));
		assert(!JSON.stringify(manifest).includes('workspace:'));
		// The manifest was packed from a copy that points at the ooxml-core tarball (not a path);
		// the tarball spec is the one `file:` reference tolerated until ooxml-core is published.
		const unpublished = new Set(['@christophervr/ooxml-core']);
		for (const [dependency, version] of Object.entries(manifest.dependencies ?? {}))
			assert(
				!String(version).includes('file:') || unpublished.has(dependency),
				`${packed.name} depends on ${dependency} through ${version}`,
			);
		assert(!JSON.stringify({ ...manifest, dependencies: undefined }).includes('file:'));
		await assertPublishedImports(path.join(target, 'dist'), packed.name);
		for (const [dependency, version] of Object.entries(manifest.peerDependencies ?? {}))
			peers.set(dependency, version);
	}
	await writeFile(
		path.join(work, 'package.json'),
		JSON.stringify({ private: true, type: 'module' }),
	);
	run(
		'npm',
		[
			'install',
			'--ignore-scripts',
			'--no-audit',
			'--no-fund',
			'--package-lock=false',
			...tarballs,
			...Array.from(peers, ([name, version]) => `${name}@${version}`),
		],
		{ cwd: work },
	);
	const packageNames = [
		'docx-viewer',
		'docx-viewer/core',
		'docx-viewer/document',
		'docx-viewer/legacy',
		'docx-viewer/web-component',
		'docx-viewer/vanilla',
		'docx-viewer/react',
		'docx-viewer/vue',
		'docx-viewer/angular',
		'docx-viewer/solid',
		'docx-core',
		'docx-core/embedded',
		'docx-legacy',
		'docx-document',
		'docx-web-component',
		'docx-bindings',
		'docx-bindings/react',
		'docx-bindings/vue',
		'docx-bindings/angular',
		'docx-bindings/solid',
	];
	const imports = packageNames.map((name) => `await import('@christophervr/${name}');`).join('\n');
	const consumer = path.join(work, 'consumer.mjs');
	await writeFile(
		consumer,
		`${imports}
import assert from 'node:assert/strict';
import * as umbrella from '@christophervr/docx-viewer';
import * as umbrellaCore from '@christophervr/docx-viewer/core';
assert.equal(typeof umbrella.createDocument, 'function');
assert.equal(typeof umbrella.mountEditor, 'function');
assert.equal(typeof umbrellaCore.loadDocx, 'function');
assert.equal(typeof globalThis.document, 'undefined', 'SSR import must not require a DOM');
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { createCollaborationAuthority } from '@christophervr/docx-web-component';
const model = createDocument();
model.blocks[0].runs = [{ text: 'Packed first line\\nSecond line', bold: true, language: 'ar-SA', rtl: false }];
model.blocks[0].direction = 'rtl';
model.blocks[0].lineSpacingTwips = 360;
model.blocks[0].lineSpacingRule = 'auto';
const bytes = await saveDocx(model);
const loaded = await loadDocument(bytes);
assert.equal(loaded.model.blocks[0].runs[0].text, 'Packed first line\\nSecond line');
assert.equal(loaded.model.blocks[0].runs[0].bold, true);
assert.equal(loaded.model.blocks[0].runs[0].language, 'ar-SA');
assert.equal(loaded.model.blocks[0].runs[0].rtl, false);
assert.equal(loaded.model.blocks[0].direction, 'rtl');
const authority = createCollaborationAuthority(model, { sessionId: 'packed-consumer' });
assert.equal(authority.currentVersion, 0);
assert.equal(authority.doc.firstChild.attrs.direction, 'rtl');
assert.equal(loaded.model.blocks[0].lineSpacingTwips, 360);
assert.equal(loaded.model.blocks[0].lineSpacingRule, 'auto');
assert.deepEqual(await loaded.save(), bytes);
loaded.model.blocks[0].lineSpacingRule = 'exact';
loaded.model.blocks[0].lineSpacingTwips = 300;
const edited = await loadDocument(await loaded.save());
assert.equal(edited.model.blocks[0].lineSpacingRule, 'exact');
assert.equal(edited.model.blocks[0].lineSpacingTwips, 300);
assert.equal(edited.model.blocks[0].runs[0].text, 'Packed first line\\nSecond line');
`,
	);
	run('node', [consumer], { cwd: work });
	const typing = path.join(work, 'consumer.ts');
	await writeFile(
		typing,
		`
import { createDocument, mountEditor, type EditorOptions } from '@christophervr/docx-viewer';
import { WordEditor as ReactEditor } from '@christophervr/docx-viewer/react';
import { WordEditor as VueEditor } from '@christophervr/docx-viewer/vue';
import { WordEditorComponent } from '@christophervr/docx-viewer/angular';
import { WordEditor as SolidEditor } from '@christophervr/docx-viewer/solid';
import SvelteEditor from '@christophervr/docx-viewer/svelte';
import { PresenceClient } from '@christophervr/docx-viewer/web-component';
import { resolveParagraphFormatting } from '@christophervr/docx-viewer/core';
const options: EditorOptions = { documentModel: createDocument(), locale: 'fr' };
void [mountEditor, options, ReactEditor, VueEditor, WordEditorComponent, SolidEditor, SvelteEditor, PresenceClient, resolveParagraphFormatting];
`,
	);
	run(
		'node',
		[
			path.join(root, 'node_modules/typescript/bin/tsc'),
			'--noEmit',
			'--strict',
			'--skipLibCheck',
			'--target',
			'ES2022',
			'--module',
			'ESNext',
			'--moduleResolution',
			'bundler',
			typing,
		],
		{ cwd: work },
	);
	// A fresh Node process must import the neutral entry without any framework installed.
	const hidden = path.join(work, 'optional-peers');
	await mkdir(hidden);
	const moved = [];
	try {
		for (const peer of ['react', 'vue', 'svelte', 'solid-js', '@angular']) {
			const source = path.join(modules, peer);
			const destination = path.join(hidden, peer.replace('@', ''));
			await rename(source, destination);
			moved.push([source, destination]);
		}
		run(
			'node',
			[
				'--input-type=module',
				'-e',
				"const pkg = await import('@christophervr/docx-viewer'); if (typeof pkg.mountEditor !== 'function' || !pkg.createDocument().blocks.length) throw new Error('Neutral entry unavailable');",
			],
			{ cwd: work },
		);
	} finally {
		for (const [source, destination] of moved) await rename(destination, source);
	}
	const bindings = JSON.parse(
		await readFile(path.join(scope, 'docx-bindings', 'package.json'), 'utf8'),
	);
	assert.equal(bindings.exports['./svelte'].svelte, './src/WordEditor.svelte');
	assert.equal(bindings.exports['./svelte'].types, './dist/svelte.d.ts');
	const svelteFile = path.join(scope, 'docx-bindings', 'src', 'WordEditor.svelte');
	const svelteSource = await readFile(svelteFile, 'utf8');
	compile(svelteSource, { filename: svelteFile, generate: 'client' });
	const viewerManifest = JSON.parse(
		await readFile(path.join(scope, 'docx-viewer', 'package.json'), 'utf8'),
	);
	assert.equal(viewerManifest.exports['./core'].import, './dist/core.js');
	assert.equal(viewerManifest.exports['./react'].types, './dist/react.d.ts');
	assert.equal(viewerManifest.exports['./svelte'].svelte, './dist/WordEditor.svelte');
	const viewerSvelte = await readFile(
		path.join(scope, 'docx-viewer', 'dist', 'WordEditor.svelte'),
		'utf8',
	);
	compile(viewerSvelte, { filename: 'WordEditor.svelte', generate: 'client' });
	// Compiling does not resolve imports, so check the copied component's value imports exist.
	const bindingExports = await import(path.join(scope, 'docx-bindings', 'dist', 'index.js'));
	for (const [, names, source] of viewerSvelte.matchAll(
		/import\s*\{([^}]*)\}\s*from\s*'([^']+)'/g,
	)) {
		assert.equal(
			source,
			'@christophervr/docx-bindings',
			'viewer WordEditor.svelte must import the bindings package',
		);
		for (const name of names
			.split(',')
			.map((part) => part.trim())
			.filter((part) => part && !part.startsWith('type ')))
			assert.equal(
				typeof bindingExports[name],
				'function',
				`bindings must export ${name} for the viewer Svelte component`,
			);
	}
	const componentBundle = await readFile(
		path.join(scope, 'docx-web-component', 'dist', 'index.js'),
		'utf8',
	);
	assert(
		componentBundle.includes('.dve-frame') && componentBundle.includes('--blue'),
		'web component CSS must be bundled as runtime text',
	);
	console.log(
		'Packed consumer imports and DOCX edit round trips succeeded; Svelte source compiles and its declaration export is present.',
	);
} finally {
	await rm(work, { recursive: true, force: true });
}
