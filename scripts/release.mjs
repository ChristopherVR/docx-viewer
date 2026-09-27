import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const mode = process.argv[2];
const requested = process.argv[3] ?? process.env.RELEASE_VERSION ?? process.env.GITHUB_REF_NAME;
if (!['--check', '--publish'].includes(mode))
	throw new Error('Usage: node scripts/release.mjs --check|--publish [vX.Y.Z]');
const version = requested?.startsWith('v') ? requested.slice(1) : requested;
if (!version || !/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(version))
	throw new Error('Provide a release tag or version such as v0.1.0.');
const releaseTag =
	process.env.RELEASE_TAG ??
	(process.env.GITHUB_REF_TYPE === 'tag' ? process.env.GITHUB_REF_NAME : undefined);
if (process.env.GITHUB_ACTIONS === 'true' && !releaseTag)
	throw new Error('Releases must run from an existing v-prefixed tag.');
if (releaseTag) {
	if (releaseTag !== `v${version}`)
		throw new Error(`Release tag ${releaseTag} does not match version ${version}.`);
	const head = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' });
	const tagHead = spawnSync('git', ['rev-parse', `refs/tags/${releaseTag}^{}`], {
		cwd: root,
		encoding: 'utf8',
	});
	if (head.status !== 0 || tagHead.status !== 0 || head.stdout.trim() !== tagHead.stdout.trim())
		throw new Error(`Checked-out HEAD does not match existing release tag ${releaseTag}.`);
}
const names = ['core', 'legacy', 'document', 'web-component', 'bindings', 'viewer'];
const packages = [];
const npm =
	process.platform === 'win32'
		? path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
		: 'npm';
const npmCommand = process.platform === 'win32' ? process.execPath : npm;
const npmPrefix = process.platform === 'win32' ? [npm] : [];
for (const directory of names) {
	const packageDir = path.join(root, 'packages', directory);
	const manifest = JSON.parse(await readFile(path.join(packageDir, 'package.json'), 'utf8'));
	if (manifest.version !== version)
		throw new Error(`${manifest.name} is ${manifest.version}, requested release is ${version}.`);
	if (manifest.private) throw new Error(`${manifest.name} is private and cannot be published.`);
	for (const dependency of Object.keys(manifest.dependencies ?? {})) {
		if (dependency.startsWith('@christophervr/docx-')) {
			const dependencyVersion = manifest.dependencies[dependency];
			if (dependencyVersion !== version)
				throw new Error(
					`${manifest.name} expects ${dependency}@${dependencyVersion}; release versions must match.`,
				);
		}
	}
	packages.push({ directory, packageDir, name: manifest.name });
}
if (mode === '--check') {
	console.log(
		`Validated ${packages.length} public packages for version ${version} in dependency order.`,
	);
	process.exit(0);
}
for (const pkg of packages) {
	const query = spawnSync(
		npmCommand,
		[...npmPrefix, 'view', `${pkg.name}@${version}`, 'version', '--json'],
		{ cwd: root, encoding: 'utf8' },
	);
	if (query.status === 0 && query.stdout.trim()) {
		console.log(`Skipping ${pkg.name}@${version}: version already exists.`);
		continue;
	}
	if (!/E404|404 Not Found|is not in this registry/i.test(`${query.stdout}\n${query.stderr}`)) {
		throw new Error(
			`Could not determine registry status for ${pkg.name}@${version}: ${query.stderr || query.stdout}`,
		);
	}
	const prerelease = version.includes('-');
	const publish = spawnSync(
		npmCommand,
		[
			...npmPrefix,
			'publish',
			'--access',
			'public',
			'--provenance',
			'--tag',
			prerelease ? 'next' : 'latest',
		],
		{
			cwd: pkg.packageDir,
			stdio: 'inherit',
			env: (() => {
				const env = { ...process.env };
				if (env.NPM_TOKEN) env.NODE_AUTH_TOKEN = env.NPM_TOKEN;
				else delete env.NODE_AUTH_TOKEN;
				return env;
			})(),
		},
	);
	if (publish.status !== 0) throw new Error(`Publishing ${pkg.name}@${version} failed.`);
}
