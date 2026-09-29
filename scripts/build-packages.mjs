import { cp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { build } from 'vite';
import { svelte } from '@sveltejs/vite-plugin-svelte';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packages = ['core', 'legacy', 'document', 'layout', 'web-component', 'bindings', 'viewer'];
const run = (command, args, options = {}) => {
	const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', ...options });
	if (result.status !== 0)
		throw new Error(
			`${command} ${args.join(' ')} failed${result.error ? `: ${result.error.message}` : ''}`,
		);
};

await rm(path.join(root, '.release-types'), { recursive: true, force: true });
run(process.execPath, [
	path.join(root, 'node_modules/typescript/bin/tsc'),
	'--project',
	'tsconfig.release.json',
]);

for (const name of packages) {
	const packageDir = path.join(root, 'packages', name);
	const distDir = path.join(packageDir, 'dist');
	await rm(distDir, { recursive: true, force: true });
	await mkdir(distDir, { recursive: true });
	const declRoot = path.join(root, '.release-types', 'packages', name, 'src');
	try {
		for (const entry of await readdir(declRoot, { withFileTypes: true })) {
			if (entry.isFile() && entry.name.endsWith('.d.ts') && !entry.name.includes('.test.')) {
				await cp(path.join(declRoot, entry.name), path.join(distDir, entry.name));
			}
		}
	} catch {
		throw new Error(`No emitted declarations found for packages/${name}/src`);
	}
	const manifest = JSON.parse(await readFile(path.join(packageDir, 'package.json'), 'utf8'));
	const inputs =
		name === 'bindings'
			? {
					index: 'src/index.ts',
					react: 'src/react.tsx',
					vue: 'src/vue.ts',
					angular: 'src/angular.ts',
					solid: 'src/solid.ts',
				}
			: name === 'viewer'
				? Object.fromEntries(
						[
							'index',
							'core',
							'document',
							'legacy',
							'web-component',
							'vanilla',
							'react',
							'vue',
							'angular',
							'svelte',
							'solid',
						].map((entry) => [entry, `src/${entry}.ts`]),
					)
				: { index: 'src/index.ts', ...(name === 'core' ? { embedded: 'src/embedded.ts' } : {}) };
	await build({
		configFile: false,
		root: packageDir,
		plugins: name === 'bindings' || name === 'web-component' ? [svelte()] : [],
		build: {
			target: 'es2022',
			outDir: distDir,
			emptyOutDir: false,
			lib: {
				entry: Object.fromEntries(
					Object.entries(inputs).map(([key, value]) => [key, path.join(packageDir, value)]),
				),
				formats: ['es'],
				fileName: (_format, entryName) => `${entryName}.js`,
			},
			rollupOptions: {
				external: [
					...Object.keys(manifest.dependencies ?? {}),
					...Object.keys(manifest.peerDependencies ?? {}),
					/^@christophervr\//,
				],
			},
		},
	});
	if (name === 'bindings') {
		await writeFile(
			path.join(distDir, 'svelte.d.ts'),
			[
				"import type { Component } from 'svelte';",
				"import type { EditorEventHandlers, EditorProps } from './index';",
				"type Props = EditorProps & { ondocumentchange?: EditorEventHandlers['document-change']; ondocumenterror?: EditorEventHandlers['document-error']; onpagechange?: EditorEventHandlers['page-change']; ondirtychange?: EditorEventHandlers['dirty-change'] };",
				'type Exports = { load(input: Uint8Array | ArrayBuffer): Promise<void>; save(): Promise<Blob>; download(fileName?: string): Promise<void>; markClean(): void; isDirty(): boolean };',
				'declare const WordEditor: Component<Props, Exports>;',
				'export default WordEditor;',
				'',
			].join('\n'),
		);
	}
}
const viewerDist = path.join(root, 'packages/viewer/dist');
// The component imports binding helpers from './index'; in the viewer package that path is the
// viewer's own entry, which does not re-export them, so point it at the bindings package instead.
const svelteSource = await readFile(
	path.join(root, 'packages/bindings/src/WordEditor.svelte'),
	'utf8',
);
const viewerSvelte = svelteSource.replace(
	/from '\.\/index'/g,
	"from '@christophervr/docx-bindings'",
);
if (viewerSvelte === svelteSource)
	throw new Error(
		"WordEditor.svelte no longer imports from './index'; update the viewer copy step",
	);
await writeFile(path.join(viewerDist, 'WordEditor.svelte'), viewerSvelte);
await writeFile(
	path.join(viewerDist, 'svelte.d.ts'),
	"export { default } from '@christophervr/docx-bindings/svelte';\n",
);
await rm(path.join(root, '.release-types'), { recursive: true, force: true });
console.log('Built publishable ESM bundles and TypeScript declarations for all six packages.');
