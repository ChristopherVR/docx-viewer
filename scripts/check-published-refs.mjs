/**
 * Guard: nothing a published package ships may import a module its manifest does not declare.
 *
 * The framework packages inline every internal workspace package (web-component, bindings); the
 * shared logic is imported from ooxml-core, which inlines ole2 itself. If bundling or
 * declaration flattening misses one, the tarball imports an unpublished package a consumer cannot
 * install. Modelled on pptx-viewer `scripts/check-published-shared-refs.mjs`, generalised from one
 * private package to "anything not in dependencies/peerDependencies".
 *
 *   node scripts/check-published-refs.mjs [core react vue angular svelte solid vanilla]
 *
 * It scans every `.js`, `.d.ts` and `.svelte` file `npm pack` would include, so it needs the
 * packages built (`bun run build:packages`).
 */
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { builtinModules } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const PUBLISHED_DIRS = ['core', 'react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];

// `from '<x>'`, bare `import '<x>'`, `import('<x>')` and `require('<x>')`; the specifier must look
// like a module id so prose such as `from": "` inside a string never matches.
const SPECIFIER =
	/(?:\bfrom\s*|\bimport\s*\(?\s*|\brequire\s*\(\s*)["']((?:@[\w.-]+\/)?[\w.-]+(?:\/[\w./-]*)?|\.{1,2}\/[\w./-]*)["']/gu;

/** Module specifiers a source file imports, de-duplicated, relative ones excluded. */
export function importedPackages(source) {
	const found = new Set();
	for (const [, specifier] of source.matchAll(SPECIFIER)) {
		if (!specifier.startsWith('.')) found.add(specifier);
	}
	return [...found];
}

const packageOf = (specifier) =>
	specifier.startsWith('@') ? specifier.split('/').slice(0, 2).join('/') : specifier.split('/')[0];

/** Specifiers in `source` that are neither Node built-ins nor declared by `manifest`. */
export function undeclaredImports(source, manifest) {
	const declared = new Set([
		...Object.keys(manifest.dependencies ?? {}),
		...Object.keys(manifest.peerDependencies ?? {}),
	]);
	return importedPackages(source).filter((specifier) => {
		const name = packageOf(specifier);
		return !(name.startsWith('node:') || builtinModules.includes(name) || declared.has(name));
	});
}

/** Manifest entries a consumer cannot install: internal packages, ole2 and local protocols. */
export function forbiddenManifestEntries(manifest) {
	const forbidden = (name) =>
		/^(?:@christophervr\/|docx-(?!core$)|ooxml-(?!core$|ui$))/u.test(name);
	const found = [];
	for (const field of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
		for (const [name, range] of Object.entries(manifest[field] ?? {})) {
			if (forbidden(name) || /^(?:workspace|file|link):/u.test(range))
				found.push(`${field}.${name}`);
		}
	}
	return found;
}

function packedFiles(dir) {
	const output = execFileSync(
		'npm',
		['pack', '--dry-run', '--ignore-scripts', '--json', '--workspaces=false'],
		{
			cwd: dir,
			encoding: 'utf8',
			stdio: ['ignore', 'pipe', 'pipe'],
			shell: process.platform === 'win32',
		},
	);
	return JSON.parse(output)[0].files.map((file) => file.path);
}

/** Scan one built package; returns the offending files and manifest entries. */
export function checkPackage(name) {
	const dir = join(ROOT, 'packages', name);
	const manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
	const files = packedFiles(dir).filter((path) => /\.(?:js|d\.ts|svelte)$/u.test(path));
	if (files.length === 0)
		throw new Error(`packages/${name} has no built files; run build:packages`);
	const offenders = [];
	for (const file of files) {
		const bad = undeclaredImports(readFileSync(join(dir, file), 'utf8'), manifest);
		if (bad.length > 0) offenders.push({ file, specifiers: bad });
	}
	return {
		name: manifest.name,
		scanned: files.length,
		offenders,
		manifest: forbiddenManifestEntries(manifest),
	};
}

function main() {
	const requested = process.argv.slice(2);
	const names = requested.length > 0 ? requested : PUBLISHED_DIRS;
	let failed = false;
	for (const name of names) {
		if (!PUBLISHED_DIRS.includes(name)) throw new Error(`Unknown package "${name}".`);
		const result = checkPackage(name);
		if (result.offenders.length === 0 && result.manifest.length === 0) {
			console.log(
				`[check-published-refs] ${result.name}: ${result.scanned} files, no undeclared imports.`,
			);
			continue;
		}
		failed = true;
		for (const { file, specifiers } of result.offenders)
			console.error(
				`[check-published-refs] ${result.name}: ${file} imports ${specifiers.join(', ')}`,
			);
		for (const entry of result.manifest)
			console.error(`[check-published-refs] ${result.name}: manifest ${entry} cannot be installed`);
	}
	if (failed) process.exitCode = 1;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
