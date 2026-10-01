#!/usr/bin/env node
/**
 * publish-released.mjs: publish the packages of a release to npm, in dependency order.
 *
 * Two modes, mirroring `.github/workflows/release.yml`:
 *   --plan release-plan.json   publish every package the plan marks `release` (the scheduled run)
 *   --tag <npm-name>@<version> re-publish exactly that existing tag (manual dispatch with `tag`)
 *
 * Authentication is npm trusted publishing (GitHub OIDC): there is no token, and `--provenance`
 * attaches the build attestation. The publisher entry for each package on npmjs.com must name this
 * repository, the workflow `release.yml` and the environment `npm`.
 *
 * Safety checks before anything is uploaded, per package:
 *   - the manifest on disk is the version being published (a re-publish checks out the tag),
 *   - no dependency range uses the `workspace:` or `file:` protocol, which no consumer can install,
 *   - no dependency names an internal workspace package or `@christophervr/ole2` (they are bundled),
 *   - every range on another package of this repo points at the version that package ships with,
 *   - a version that already exists on the registry is skipped, so a re-run is idempotent,
 *   - a version older than the registry's `latest` is published under the `old` dist-tag so it
 *     cannot move `latest` backwards.
 *
 * `--dry-run` runs all the checks and prints the `npm publish` commands without running them.
 * Node >= 22, npm >= 11.5.1 (trusted publishing) on the publishing runner.
 */

import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { forbiddenManifestEntries } from './check-published-refs.mjs';
import { cmpSemver, PACKAGES } from './release-plan.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies', 'devDependencies'];
const npm = (args, options = {}) =>
	spawnSync('npm', args, {
		cwd: ROOT,
		encoding: 'utf8',
		shell: process.platform === 'win32',
		...options,
	});

/** `{ key, npm, dir, version }` targets, in dependency order, from a plan or a tag. */
export function resolveTargets({ plan, tag }) {
	if (tag) {
		if (/[^A-Za-z0-9.@/_-]/u.test(tag) || !tag.includes('@', 1)) {
			throw new Error(`Invalid tag '${tag}'. Expected <npm-name>@<semver>.`);
		}
		const at = tag.lastIndexOf('@');
		const [name, version] = [tag.slice(0, at), tag.slice(at + 1)];
		const key = Object.keys(PACKAGES).find((k) => PACKAGES[k].npm === name);
		if (!key) throw new Error(`Unknown package '${name}' (from tag ${tag}).`);
		return [{ key, npm: name, dir: PACKAGES[key].dir, version }];
	}
	return plan.order
		.filter((key) => plan.packages[key].release)
		.map((key) => ({ key, ...pick(plan.packages[key]) }));
}
const pick = ({ npm: name, dir, version }) => ({ npm: name, dir, version });

/** Throws when the manifest on disk is not publishable as `target`. */
export function verifyManifest(target, versions = workspaceVersions()) {
	const manifest = JSON.parse(readFileSync(join(ROOT, target.dir, 'package.json'), 'utf8'));
	if (manifest.name !== target.npm)
		throw new Error(`${target.dir} is ${manifest.name}, not ${target.npm}.`);
	if (manifest.version !== target.version) {
		throw new Error(`${target.npm} is ${manifest.version} on disk, release is ${target.version}.`);
	}
	if (manifest.private) throw new Error(`${target.npm} is private and cannot be published.`);
	// Internal packages and ole2 are inlined at build time; a manifest naming one is uninstallable.
	const forbidden = forbiddenManifestEntries(manifest);
	if (forbidden.length > 0)
		throw new Error(`${target.npm} depends on unpublished packages: ${forbidden.join(', ')}.`);
	for (const field of FIELDS) {
		for (const [dep, range] of Object.entries(manifest[field] ?? {})) {
			if (/^(?:workspace|file|link):/u.test(range)) {
				throw new Error(
					`${target.npm}: ${field}["${dep}"] is "${range}", which cannot be installed.`,
				);
			}
			const sibling = versions.get(dep);
			if (field !== 'devDependencies' && sibling && range.replace(/^[\^~]/u, '') !== sibling) {
				throw new Error(`${target.npm}: ${dep} is "${range}" but that package is at ${sibling}.`);
			}
		}
	}
}

function workspaceVersions() {
	return new Map(
		Object.values(PACKAGES).map((meta) => {
			const manifest = JSON.parse(readFileSync(join(ROOT, meta.dir, 'package.json'), 'utf8'));
			return [meta.npm, manifest.version];
		}),
	);
}

/** Registry state of `name@version`: 'exists', 'missing', or throws on an unexpected error. */
function registryState(name, version) {
	const result = npm(['view', `${name}@${version}`, 'version']);
	if (result.status === 0 && result.stdout.trim()) return 'exists';
	if (/E404|404 Not Found|is not in this registry/iu.test(`${result.stdout}${result.stderr}`)) {
		return 'missing';
	}
	throw new Error(`Could not query npm for ${name}@${version}: ${result.stderr || result.stdout}`);
}

function distTag(name, version) {
	const latest = npm(['view', name, 'version']);
	const current = latest.status === 0 ? latest.stdout.trim() : '';
	return /^\d+\.\d+\.\d+$/u.test(current) && cmpSemver(version, current) < 0 ? 'old' : 'latest';
}

function main() {
	const argv = process.argv.slice(2);
	const value = (flag) => (argv.includes(flag) ? argv[argv.indexOf(flag) + 1] : undefined);
	const dryRun = argv.includes('--dry-run');
	const targets = resolveTargets(
		argv.includes('--tag')
			? { tag: value('--tag') }
			: {
					plan: JSON.parse(
						readFileSync(value('--plan') ?? join(ROOT, 'release-plan.json'), 'utf8'),
					),
				},
	);
	if (targets.length === 0) {
		console.log('Nothing to publish.');
		return;
	}
	for (const target of targets) {
		console.log(`--- ${target.npm}@${target.version} ---`);
		verifyManifest(target);
		if (registryState(target.npm, target.version) === 'exists') {
			console.log('already published, skipping');
			continue;
		}
		const args = [
			'publish',
			'--provenance',
			'--access',
			'public',
			'--tag',
			distTag(target.npm, target.version),
		];
		if (dryRun) {
			console.log(`dry run: (cd ${target.dir} && npm ${args.join(' ')})`);
			continue;
		}
		const result = npm(args, { cwd: join(ROOT, target.dir), stdio: 'inherit' });
		if (result.status !== 0) throw new Error(`Publishing ${target.npm}@${target.version} failed.`);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
