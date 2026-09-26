import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
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
	for (const name of ['core', 'legacy', 'document', 'web-component', 'bindings']) {
		const packageDir = path.join(root, 'packages', name);
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
		assert(!JSON.stringify(manifest).includes('file:'));
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
		'docx-core',
		'docx-core/embedded',
		'docx-legacy',
		'docx-document',
		'docx-web-component',
		'docx-bindings',
		'docx-bindings/react',
		'docx-bindings/vue',
		'docx-bindings/angular',
	];
	const imports = packageNames.map((name) => `await import('@christophervr/${name}');`).join('\n');
	const consumer = path.join(work, 'consumer.mjs');
	await writeFile(
		consumer,
		`${imports}
import assert from 'node:assert/strict';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
const model = createDocument();
model.blocks[0].runs = [{ text: 'Packed first line\\nSecond line', bold: true }];
model.blocks[0].lineSpacingTwips = 360;
model.blocks[0].lineSpacingRule = 'auto';
const bytes = await saveDocx(model);
const loaded = await loadDocument(bytes);
assert.equal(loaded.model.blocks[0].runs[0].text, 'Packed first line\\nSecond line');
assert.equal(loaded.model.blocks[0].runs[0].bold, true);
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
	const bindings = JSON.parse(
		await readFile(path.join(scope, 'docx-bindings', 'package.json'), 'utf8'),
	);
	assert.equal(bindings.exports['./svelte'].svelte, './src/WordEditor.svelte');
	assert.equal(bindings.exports['./svelte'].types, './dist/svelte.d.ts');
	const svelteFile = path.join(scope, 'docx-bindings', 'src', 'WordEditor.svelte');
	const svelteSource = await readFile(svelteFile, 'utf8');
	compile(svelteSource, { filename: svelteFile, generate: 'client' });
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
