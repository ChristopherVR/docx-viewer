import { cp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { build } from 'vite';
import { rollup } from 'rollup';
import { dts } from 'rollup-plugin-dts';

/**
 * Builds the published packages: `core` (a thin entry over ooxml-core/docx) and
 * one package per framework. The internal workspace packages (web-component, bindings) are
 * `private` and are inlined into each framework bundle. Layout, document loading and the legacy
 * .doc reader live in ooxml-core (which inlines the ole2 codecs itself), so a
 * tarball only imports `docx-core`, `ooxml-core`, the prosemirror libraries and its framework
 * peers. Declarations are flattened the same way.
 */
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packagesDir = path.join(root, 'packages');
const typesDir = path.join(root, '.release-types');

/** Published packages and the source entry behind each emitted `dist/<name>.js`. */
export const PUBLISHED = {
	core: { index: 'src/index.ts', embedded: 'src/embedded.ts' },
	react: { index: 'src/index.ts' },
	vue: { index: 'src/index.ts' },
	angular: { index: 'src/index.ts' },
	solid: { index: 'src/index.ts' },
	vanilla: { index: 'src/index.ts' },
	svelte: { runtime: 'src/runtime.ts' },
};

/** Private workspace packages that get inlined, by import specifier -> source file. */
const BINDING_ENTRIES = {
	react: 'react.tsx',
	vue: 'vue.ts',
	angular: 'angular.ts',
	solid: 'solid.ts',
	common: 'common.ts',
};
export const INTERNAL_SOURCES = new Map([
	['docx-web-component', 'packages/web-component/src/index.ts'],
	['docx-bindings', 'packages/bindings/src/index.ts'],
	...Object.entries(BINDING_ENTRIES).map(([name, file]) => [
		`docx-bindings/${name}`,
		`packages/bindings/src/${file}`,
	]),
]);

const run = (command, args) => {
	const result = spawnSync(command, args, { cwd: root, stdio: 'inherit' });
	if (result.status !== 0) throw new Error(`${command} ${args.join(' ')} failed`);
};
const readManifest = async (name) =>
	JSON.parse(await readFile(path.join(packagesDir, name, 'package.json'), 'utf8'));
const isDependency = (names) => (id) =>
	names.some((name) => id === name || id.startsWith(`${name}/`));
const dependencyNames = (manifest) => [
	...Object.keys(manifest.dependencies ?? {}),
	...Object.keys(manifest.peerDependencies ?? {}),
];

async function buildJavaScript(name, manifest, distDir) {
	await build({
		configFile: false,
		root: path.join(packagesDir, name),
		logLevel: 'warn',
		resolve: {
			alias: [...INTERNAL_SOURCES].map(([find, file]) => ({
				find: new RegExp(`^${find}$`),
				replacement: path.join(root, file),
			})),
		},
		build: {
			target: 'es2022',
			outDir: distDir,
			emptyOutDir: false,
			lib: {
				entry: Object.fromEntries(
					Object.entries(PUBLISHED[name]).map(([key, file]) => [
						key,
						path.join(packagesDir, name, file),
					]),
				),
				formats: ['es'],
				fileName: (_format, entryName) => `${entryName}.js`,
			},
			rollupOptions: { external: isDependency(dependencyNames(manifest)) },
		},
	});
}

/** Inlines internal declarations into one flat `.d.ts` per entry; dependencies stay imports. */
async function bundleDeclarations(name, manifest, distDir) {
	const internal = new Map(
		[...INTERNAL_SOURCES].map(([id, file]) => [
			id,
			path.join(typesDir, file.replace(/\.tsx?$/, '.d.ts')),
		]),
	);
	const isExternal = isDependency(dependencyNames(manifest));
	for (const key of Object.keys(PUBLISHED[name])) {
		const bundle = await rollup({
			input: path.join(typesDir, 'packages', name, 'src', `${key}.d.ts`),
			external: (id) =>
				!internal.has(id) && !id.startsWith('.') && !path.isAbsolute(id) && isExternal(id),
			onwarn: (warning, warn) => {
				if (warning.code !== 'UNUSED_EXTERNAL_IMPORT') warn(warning);
			},
			plugins: [
				{ name: 'internal-declarations', resolveId: (id) => internal.get(id) ?? null },
				dts(),
			],
		});
		await bundle.write({ file: path.join(distDir, `${key}.d.ts`), format: 'es' });
		await bundle.close();
	}
}

/** Svelte ships the component source; its helper import is repointed at the bundled runtime. */
async function writeSvelteComponent(distDir) {
	const source = await readFile(path.join(packagesDir, 'bindings/src/WordEditor.svelte'), 'utf8');
	const component = source.replace(/from '\.\/index'/g, "from './runtime.js'");
	if (component === source)
		throw new Error("WordEditor.svelte no longer imports from './index'; update this step");
	await writeFile(path.join(distDir, 'WordEditor.svelte'), component);
	await writeFile(
		path.join(distDir, 'index.d.ts'),
		[
			"import type { Component } from 'svelte';",
			"import type { EditorEventHandlers, EditorProps } from './runtime';",
			"type Props = EditorProps & { ondocumentchange?: EditorEventHandlers['document-change']; ondocumenterror?: EditorEventHandlers['document-error']; onpagechange?: EditorEventHandlers['page-change']; ondirtychange?: EditorEventHandlers['dirty-change'] };",
			'type Exports = { load(input: Uint8Array | ArrayBuffer): Promise<void>; save(): Promise<Blob>; download(fileName?: string): Promise<void>; markClean(): void; isDirty(): boolean };',
			'declare const WordEditor: Component<Props, Exports>;',
			'export default WordEditor;',
			'',
		].join('\n'),
	);
}

async function copyCoreDeclarations(distDir) {
	const from = path.join(typesDir, 'packages', 'core', 'src');
	for (const entry of await readdir(from, { withFileTypes: true })) {
		if (entry.isFile() && entry.name.endsWith('.d.ts') && !entry.name.includes('.test.'))
			await cp(path.join(from, entry.name), path.join(distDir, entry.name));
	}
}

async function main() {
	await rm(typesDir, { recursive: true, force: true });
	run(process.execPath, [
		path.join(root, 'node_modules/typescript/bin/tsc'),
		'--project',
		'tsconfig.release.json',
	]);
	// core first: the framework packages leave docx-core as an import.
	for (const name of Object.keys(PUBLISHED)) {
		const distDir = path.join(packagesDir, name, 'dist');
		await rm(distDir, { recursive: true, force: true });
		await mkdir(distDir, { recursive: true });
		const manifest = await readManifest(name);
		await buildJavaScript(name, manifest, distDir);
		if (name === 'core') await copyCoreDeclarations(distDir);
		else await bundleDeclarations(name, manifest, distDir);
		if (name === 'svelte') await writeSvelteComponent(distDir);
	}
	await rm(typesDir, { recursive: true, force: true });
	console.log(
		`Built ${Object.keys(PUBLISHED).length} publishable packages (docx-core and six framework packages).`,
	);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) await main();
