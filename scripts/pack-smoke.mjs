import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { compile } from 'svelte/compiler';
import { forbiddenManifestEntries, undeclaredImports } from './check-published-refs.mjs';

/**
 * Packs the seven published packages (docx-core and the six self-contained framework packages),
 * installs the tarballs together into a clean consumer, and exercises them: every entry imports in
 * Node without a DOM, DOCX and legacy .doc files load through each framework package, and no
 * tarball imports an internal workspace package or `@christophervr/ole2`.
 */
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

const FRAMEWORKS = ['react', 'vue', 'angular', 'solid', 'svelte', 'vanilla'];
const PUBLISHED = ['core', ...FRAMEWORKS];
const entryOf = (name) => (name === 'svelte' ? 'dist/runtime.js' : 'dist/index.js');
const specifierOf = (name) =>
	name === 'core' ? 'docx-core' : `docx-${name}-viewer${name === 'svelte' ? '/runtime' : ''}`;

async function files(directory) {
	const result = [];
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const target = path.join(directory, entry.name);
		if (entry.isDirectory()) result.push(...(await files(target)));
		else result.push(target);
	}
	return result;
}

/** Inspects the extracted tarball: manifest, entry files and every import it ships. */
async function inspectTarball(name, packed, directory) {
	const manifest = JSON.parse(await readFile(path.join(directory, 'package.json'), 'utf8'));
	assert.equal(manifest.private, undefined, `${packed.name} must not be private`);
	assert(!JSON.stringify(manifest).includes('workspace:'), `${packed.name} has a workspace: range`);
	assert(!JSON.stringify(manifest).includes('file:'), `${packed.name} has a file: range`);
	assert.deepEqual(forbiddenManifestEntries(manifest), [], `${packed.name} manifest`);
	assert(
		packed.files.some((entry) => entry.path === entryOf(name)),
		`${packed.name} has no built JavaScript entry`,
	);
	assert(
		packed.files.some((entry) => entry.path === 'dist/index.d.ts'),
		`${packed.name} has no TypeScript declarations`,
	);
	if (name !== 'core') {
		assert.deepEqual(
			Object.keys(manifest.dependencies).filter((dep) => dep.startsWith('@christophervr/')),
			['@christophervr/docx-core', '@christophervr/office-ui', '@christophervr/ooxml-core'],
			`${packed.name} may depend on no project package but docx-core, office-ui and ooxml-core`,
		);
	}
	for (const file of await files(path.join(directory, 'dist'))) {
		if (!/\.(?:js|d\.ts|svelte)$/.test(file)) continue;
		assert.deepEqual(
			undeclaredImports(await readFile(file, 'utf8'), manifest),
			[],
			`${packed.name} imports an undeclared or unpublished module in ${path.relative(directory, file)}`,
		);
	}
	return manifest;
}

const consumerSource = `import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
${FRAMEWORKS.map((name) => `import * as ${name} from '@christophervr/${specifierOf(name)}';`).join('\n')}
import * as core from '@christophervr/docx-core';
import * as embedded from '@christophervr/docx-core/embedded';
assert.equal(typeof globalThis.document, 'undefined', 'SSR import must not require a DOM');
assert.equal(typeof core.createDocument, 'function');
assert.equal(typeof embedded, 'object');
const frameworks = { react, vue, angular, solid, svelte, vanilla };
const component = { react: 'WordEditor', vue: 'WordEditor', angular: 'WordEditorComponent', solid: 'WordEditor' };
for (const [name, entry] of Object.entries(frameworks)) {
	if (component[name]) assert(entry[component[name]], name + ' must export its component');
	for (const helper of ['loadDocument', 'detectDocumentFormat', 'registerDocxEditor', 'normalizeEditorLocale', 'createDocument'])
		assert.equal(typeof entry[helper], 'function', name + ' must export ' + helper);
}
assert.equal(typeof vanilla.mountEditor, 'function');
assert.equal(typeof svelte.mountEditor, 'function');

const model = core.createDocument();
model.blocks[0].runs = [{ text: 'Packed first line\\nSecond line', bold: true, language: 'ar-SA', rtl: false }];
model.blocks[0].direction = 'rtl';
model.blocks[0].lineSpacingTwips = 360;
model.blocks[0].lineSpacingRule = 'auto';
const bytes = await core.saveDocx(model);
const doc = new Uint8Array(await readFile('fixtures/ole-word-97.doc'));
for (const [name, entry] of Object.entries(frameworks)) {
	// DOCX through each package's bundled loader, on the shared docx-core.
	assert.equal(entry.detectDocumentFormat(bytes), 'docx');
	const loaded = await entry.loadDocument(bytes);
	assert.equal(loaded.model.blocks[0].runs[0].text, 'Packed first line\\nSecond line', name);
	assert.equal(loaded.model.blocks[0].runs[0].bold, true, name);
	assert.equal(loaded.model.blocks[0].direction, 'rtl', name);
	assert.deepEqual(await loaded.save(), bytes, name + ' no-op save must return the original bytes');
	loaded.model.blocks[0].lineSpacingRule = 'exact';
	loaded.model.blocks[0].lineSpacingTwips = 300;
	const edited = await entry.loadDocument(await loaded.save());
	assert.equal(edited.model.blocks[0].lineSpacingRule, 'exact', name);
	// Legacy .doc through ooxml-core/docx/load, which inlines the ole2 codecs: nothing depends on ole2.
	assert.equal(entry.detectDocumentFormat(doc), 'doc');
	const legacy = await entry.loadDocument(doc);
	assert(legacy.model.blocks.length > 0 && legacy.model.blocks[0].type === 'paragraph', name + ' legacy .doc');
	assert.match(legacy.model.warnings[0], /main-body text only/, name);
	assert.deepEqual(await legacy.save(), doc, name + ' legacy no-op save must return the original bytes');
}
const authority = vanilla.createCollaborationAuthority(model, { sessionId: 'packed-consumer' });
assert.equal(authority.currentVersion, 0);
assert.equal(authority.doc.firstChild.attrs.direction, 'rtl');
`;

const typingSource = `import { createDocument, resolveParagraphFormatting, type DocumentModel } from '@christophervr/docx-core';
import { WordEditor as ReactEditor, loadDocument, type EditorOptions } from '@christophervr/docx-react-viewer';
import { WordEditor as VueEditor } from '@christophervr/docx-vue-viewer';
import { WordEditorComponent } from '@christophervr/docx-angular-viewer';
import { WordEditor as SolidEditor } from '@christophervr/docx-solid-viewer';
import SvelteEditor from '@christophervr/docx-svelte-viewer';
import { mountEditor, PresenceClient, type EditorHandle } from '@christophervr/docx-vanilla-viewer';
const options: EditorOptions = { documentModel: createDocument(), locale: 'fr' };
const model: DocumentModel = options.documentModel!;
void [model, ReactEditor, loadDocument, VueEditor, WordEditorComponent, SolidEditor, SvelteEditor, mountEditor, PresenceClient, resolveParagraphFormatting];
export type Handle = EditorHandle;
`;

/** The Svelte component compiles and every helper it imports exists in the bundled runtime. */
async function checkSvelte(inspected, installed) {
	const manifest = JSON.parse(
		await readFile(path.join(inspected, 'svelte', 'package.json'), 'utf8'),
	);
	assert.equal(manifest.exports['.'].svelte, './dist/WordEditor.svelte');
	assert.equal(manifest.exports['./runtime'].import, './dist/runtime.js');
	const source = await readFile(
		path.join(inspected, 'svelte', 'dist', 'WordEditor.svelte'),
		'utf8',
	);
	compile(source, { filename: 'WordEditor.svelte', generate: 'client' });
	const runtime = await import(
		pathToFileURL(path.join(installed, 'docx-svelte-viewer', 'dist', 'runtime.js')).href
	);
	for (const [, names, from] of source.matchAll(/import\s*\{([^}]*)\}\s*from\s*'([^']+)'/g)) {
		assert.equal(from, './runtime.js', 'WordEditor.svelte must import its sibling runtime');
		for (const name of names
			.split(',')
			.map((part) => part.trim())
			.filter((part) => part && !part.startsWith('type ')))
			assert.equal(typeof runtime[name], 'function', `runtime.js must export ${name}`);
	}
}

try {
	const tarballs = [];
	const peers = new Map();
	const inspected = path.join(work, 'inspect');
	for (const name of PUBLISHED) {
		const packed = JSON.parse(
			run('npm', ['pack', '--json', '--pack-destination', work, path.join(root, 'packages', name)]),
		)[0];
		tarballs.push(path.join(work, packed.filename));
		const target = path.join(inspected, name);
		await mkdir(target, { recursive: true });
		// Relative paths: GNU tar (Git for Windows) reads `C:` as a remote host.
		run('tar', ['-xzf', packed.filename, '--strip-components=1', '-C', `inspect/${name}`], {
			cwd: work,
		});
		const manifest = await inspectTarball(name, packed, target);
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
	const installed = path.join(work, 'node_modules', '@christophervr');
	assert.deepEqual(
		(await readdir(installed)).filter((name) =>
			/ole2|docx-(?:legacy|document|layout|bindings|web-component|viewer)$/.test(name),
		),
		[],
		'no internal package or ole2 may be installed alongside the published ones',
	);

	await mkdir(path.join(work, 'fixtures'));
	await writeFile(
		path.join(work, 'fixtures', 'ole-word-97.doc'),
		await readFile(path.join(root, 'tests/support/ole-word-97.doc')),
	);
	await writeFile(path.join(work, 'consumer.mjs'), consumerSource);
	run('node', [path.join(work, 'consumer.mjs')], { cwd: work });
	await writeFile(path.join(work, 'consumer.ts'), typingSource);
	run(
		'node',
		[
			path.join(root, 'node_modules/typescript/bin/tsc'),
			...['--noEmit', '--strict', '--skipLibCheck', '--target', 'ES2022', '--module', 'ESNext'],
			...['--moduleResolution', 'bundler', '--jsx', 'react-jsx', '--experimentalDecorators'],
			path.join(work, 'consumer.ts'),
		],
		{ cwd: work },
	);
	await checkSvelte(inspected, installed);
	for (const name of FRAMEWORKS) {
		const bundle = await readFile(path.join(inspected, name, entryOf(name)), 'utf8');
		assert(
			bundle.includes('.dve-frame') && bundle.includes('--blue'),
			`${name}: web component CSS must be bundled as runtime text`,
		);
	}
	console.log(
		'Packed consumer imports, DOCX and legacy .doc loading, typings and bundle checks succeeded for docx-core and all six framework packages.',
	);
} finally {
	await rm(work, { recursive: true, force: true });
}
