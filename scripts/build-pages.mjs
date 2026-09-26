import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, rmSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const docs = resolve(root, 'docs');
const dist = resolve(docs, '.vitepress', 'dist');
const node = process.execPath;
const vitepressCli = resolve(docs, 'node_modules', 'vitepress', 'bin', 'vitepress.js');
const viteCli = resolve(root, 'node_modules', 'vite', 'bin', 'vite.js');
const demos = [
	{ framework: 'react', route: 'demo' },
	{ framework: 'vue', route: 'demo-vue' },
	{ framework: 'angular', route: 'demo-angular' },
	{ framework: 'vanilla', route: 'demo-vanilla' },
	{ framework: 'svelte', route: 'demo-svelte' },
];

function run(script, args, env = process.env) {
	const result = spawnSync(node, [script, ...args], { cwd: root, env, stdio: 'inherit' });
	if (result.error) throw result.error;
	if (result.status !== 0)
		throw new Error(`Command failed (${result.status}): ${script} ${args.join(' ')}`);
}

if (!existsSync(resolve(docs, 'node_modules', 'vitepress'))) {
	throw new Error(
		'VitePress is missing. Install the docs dependencies with `bun install --cwd docs`.',
	);
}

run(vitepressCli, ['build', 'docs']);

for (const { framework, route } of demos) {
	const outDir = resolve(dist, route);
	if (existsSync(outDir)) rmSync(outDir, { recursive: true, force: true });
	mkdirSync(outDir, { recursive: true });
	run(
		viteCli,
		[
			'build',
			'demos/demo-vanilla',
			'--config',
			'vite.config.ts',
			'--base',
			`/docx-viewer/${route}/`,
			'--outDir',
			outDir,
			'--emptyOutDir',
		],
		{
			...process.env,
			DEMO_BASE: `/docx-viewer/${route}/`,
			VITE_DEMO_FRAMEWORK: framework,
		},
	);
	const index = resolve(outDir, 'index.html');
	if (!existsSync(index) || statSync(index).size === 0) {
		throw new Error(`${framework} demo build did not produce ${index}`);
	}
}

console.log('Built VitePress documentation and all five framework demos.');
