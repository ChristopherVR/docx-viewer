import { cp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { build } from 'vite';
import { svelte } from '@sveltejs/vite-plugin-svelte';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packages = ['core', 'legacy', 'document', 'web-component', 'bindings', 'viewer'];
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
			"import type { Component } from 'svelte';\nimport type { DocumentModel } from '@christophervr/docx-core';\ndeclare const WordEditor: Component<{ documentModel?: DocumentModel; readOnly?: boolean; locale?: string; ondocumentchange?: (model: DocumentModel) => void; ondocumenterror?: (error: Error) => void; }>;\nexport default WordEditor;\n",
		);
	}
}
const viewerDist = path.join(root, 'packages/viewer/dist');
await cp(
	path.join(root, 'packages/bindings/src/WordEditor.svelte'),
	path.join(viewerDist, 'WordEditor.svelte'),
);
await writeFile(
	path.join(viewerDist, 'svelte.d.ts'),
	"export { default } from '@christophervr/docx-bindings/svelte';\n",
);
await rm(path.join(root, '.release-types'), { recursive: true, force: true });
console.log('Built publishable ESM bundles and TypeScript declarations for all six packages.');
