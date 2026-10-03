#!/usr/bin/env node
/**
 * release-plan.mjs: decide which packages to version and publish, with an INDEPENDENT version
 * line per package.
 *
 * Adapted from ChristopherVR/pptx-viewer `scripts/release-plan.mjs` (Apache-2.0). The planning
 * model is the same; differences: dependencies between packages are read from the manifests
 * instead of a hand-kept `triggers` list, internal dependency ranges are rewritten when a
 * dependency is released, and a package that was never published is released at its manifest
 * version instead of being bumped. Candidate for centralisation (the config block is the only
 * repo-specific part): the same file is vendored in ooxml-core.
 *
 * Every published package carries its own version and its own git tag `<npm-name>@<version>`
 * (e.g. `docx-core@0.2.0`). From the history since each package's last tag this
 * script computes which packages changed (directly, or because an internal dependency is being
 * released) and the next version for each. The bump level follows Conventional Commits: a
 * breaking change (`!` or a BREAKING CHANGE footer) bumps major, `feat` bumps minor, anything
 * else patch. Unchanged packages get no tag, no GitHub release and no npm publish.
 *
 *   node scripts/release-plan.mjs             # print the plan (dry run, writes release-plan.json)
 *   node scripts/release-plan.mjs --no-npm    # skip npm lookups (offline)
 *   node scripts/release-plan.mjs --write     # also stamp versions and internal ranges
 *
 * It is the single source of truth for `.github/workflows/release.yml`; the rich per-package
 * detail is read back from `release-plan.json`. Under GitHub Actions it appends `any_changed` to
 * `$GITHUB_OUTPUT`.
 */

import { execFileSync } from 'node:child_process';
import { appendFileSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Internal, `private` workspace packages. They are never published: their code is bundled into
 * every framework package. A change in either therefore changes what every framework package
 * ships. (Layout, document loading and legacy .doc support live in `ooxml-core`, a
 * regular dependency of the framework packages, so a core bump shows up in their manifests.)
 */
export const INTERNAL_DIRS = ['packages/web-component', 'packages/bindings'];

/**
 * Publishable packages, in the order the planner reports them. `dir` is the source directory
 * and `npm` the published name. Dependencies between these packages are NOT listed here: any
 * `dependencies` / `peerDependencies` / `optionalDependencies` entry naming another package of
 * this table is an internal dependency, and releasing it re-releases the dependent (every
 * framework package depends on `docx-core`, so a core release releases all of them).
 * Registry dependencies absent from this table (`ooxml-core`, and
 * `ooxml-ui` in every framework package) are external: a range change in a
 * manifest releases that package only, as for any third-party dependency.
 * `triggers` (optional) are other directories whose published files also force a release of this
 * package, because their code is inlined into it. Same shape as pptx-viewer's SHARED_DIR trigger.
 * `paths` (optional) narrows what counts as a published file, for a package that is the repo
 * root: entries ending in `/` are directories, anything else a single file.
 */
export const PACKAGES = {
	mcp: { dir: 'mcp', npm: 'docx-viewer-mcp' },
	core: { dir: 'packages/core', npm: 'docx-core' },
	react: {
		dir: 'packages/react',
		npm: 'docx-react-viewer',
		triggers: INTERNAL_DIRS,
	},
	vue: { dir: 'packages/vue', npm: 'docx-vue-viewer', triggers: INTERNAL_DIRS },
	angular: {
		dir: 'packages/angular',
		npm: 'docx-angular-viewer',
		triggers: INTERNAL_DIRS,
	},
	svelte: {
		dir: 'packages/svelte',
		npm: 'docx-svelte-viewer',
		triggers: INTERNAL_DIRS,
	},
	solid: {
		dir: 'packages/solid',
		npm: 'docx-solid-viewer',
		triggers: INTERNAL_DIRS,
	},
	vanilla: {
		dir: 'packages/vanilla',
		npm: 'docx-vanilla-viewer',
		triggers: INTERNAL_DIRS,
	},
};

/**
 * Paths outside any package directory that still change what every published artifact contains,
 * so a change forces a re-release of all of them: the shared build pipeline.
 */
export const GLOBAL_TRIGGERS = ['scripts/build-packages.mjs', 'tsconfig.release.json'];

/** package.json fields whose change never alters what a consumer receives. */
const IGNORED_MANIFEST_FIELDS = ['version', 'scripts', 'devDependencies'];
const DEP_FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies'];
const BUMP_RANK = { patch: 0, minor: 1, major: 2 };
const SEMVER = /^\d+\.\d+\.\d+$/u;

export function cmpSemver(a, b) {
	const pa = a.split('.').map(Number);
	const pb = b.split('.').map(Number);
	for (let i = 0; i < 3; i++) {
		const d = (pa[i] || 0) - (pb[i] || 0);
		if (d !== 0) return d > 0 ? 1 : -1;
	}
	return 0;
}

export const maxSemver = (versions) =>
	versions.reduce((best, v) => (cmpSemver(v, best) > 0 ? v : best), '0.0.0');

export function bumpVersion(version, level) {
	const [maj, min, patch] = version.split('.').map(Number);
	if (level === 'major') return `${maj + 1}.0.0`;
	if (level === 'minor') return `${maj}.${min + 1}.0`;
	return `${maj}.${min}.${patch + 1}`;
}

/** Conventional Commit bump level of one commit message. */
export function commitLevel(subject, body = '') {
	if (/^[a-z]+(?:\([^)]*\))?!:/iu.test(subject) || /(?:^|\n)BREAKING[ -]CHANGE:/u.test(body)) {
		return 'major';
	}
	return /^feat(?:\([^)]*\))?:/u.test(subject) ? 'minor' : 'patch';
}

/** Files that ship in a published artifact (exclude tests). */
export function isPublishedFile(path) {
	if (/\.(?:test|spec)\.[cm]?[jt]sx?$/u.test(path)) return false;
	return !path.includes('/__tests__/') && !path.includes('/e2e/');
}

/** JSON text with object keys sorted, so reordering a manifest is not a change. */
const canonical = (value) =>
	JSON.stringify(value, (_key, v) =>
		v && typeof v === 'object' && !Array.isArray(v)
			? Object.fromEntries(Object.entries(v).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
			: v,
	);

const readJson = (path) => JSON.parse(readFileSync(path, 'utf8'));
const writeJson = (path, data) => writeFileSync(path, `${JSON.stringify(data, null, '\t')}\n`);

/** `npm view <name> version`: the version, or null when the package does not exist (E404). */
export function npmVersion(name) {
	try {
		const out = execFileSync('npm', ['view', name, 'version'], {
			encoding: 'utf8',
			shell: process.platform === 'win32',
			stdio: ['ignore', 'pipe', 'pipe'],
		});
		return out.trim() || null;
	} catch (error) {
		if (/E404|404 Not Found|is not in this registry/iu.test(String(error.stderr ?? ''))) {
			return null;
		}
		throw new Error(`Could not query npm for ${name}: ${error.stderr || error.message}`);
	}
}

/**
 * Compute the release plan.
 * @param {object} options
 * @param {string} options.root repository root (a git work tree)
 * @param {Record<string, {dir: string, npm: string, paths?: string[]}>} options.packages
 * @param {string[]} [options.globalTriggers]
 * @param {(name: string) => string | null} [options.npm] registry lookup; omit for offline
 */
export function planRelease({ root, packages: table, globalTriggers = [], npm }) {
	const git = (args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
	const manifestPath = (meta) => join(root, meta.dir, 'package.json');
	const names = new Set(Object.values(table).map((m) => m.npm));
	const scopeOf = (meta) => meta.paths ?? [meta.dir];
	const triggersOf = (meta) => meta.triggers ?? [];
	/** Own files plus the internal directories bundled into the package. */
	const fullScopeOf = (meta) => [...scopeOf(meta), ...triggersOf(meta)];
	const under = (file, target) =>
		file === target || file.startsWith(target.endsWith('/') ? target : `${target}/`);
	const touches = (files, targets) => files.some((f) => targets.some((t) => under(f, t)));

	/** Internal dependencies (package keys) of one package, from its manifest at HEAD. */
	const internalDeps = (meta) => {
		const manifest = readJson(manifestPath(meta));
		const found = new Set();
		for (const field of DEP_FIELDS) {
			for (const dep of Object.keys(manifest[field] ?? {})) {
				const key = Object.keys(table).find((k) => table[k].npm === dep);
				if (key) found.add(key);
			}
		}
		return [...found];
	};

	/** Keys in dependency order: every package comes after the packages it depends on. */
	const order = [];
	const visit = (key, trail = []) => {
		if (order.includes(key)) return;
		if (trail.includes(key)) throw new Error(`Dependency cycle: ${[...trail, key].join(' -> ')}`);
		for (const dep of internalDeps(table[key])) visit(dep, [...trail, key]);
		order.push(key);
	};
	Object.keys(table).forEach((key) => visit(key));

	const baselineTag = (name) => {
		const tags = git(['tag', '--list', `${name}@*`, '--sort=-version:refname'])
			.split('\n')
			.map((t) => t.trim())
			.filter((t) => t.startsWith(`${name}@`) && SEMVER.test(t.slice(name.length + 1)));
		for (const tag of tags) {
			try {
				// A tag on HEAD is a valid baseline: it yields an empty diff (no re-release).
				git(['merge-base', '--is-ancestor', tag, 'HEAD']);
				return tag;
			} catch {
				// Tag on another line of history; keep looking.
			}
		}
		return null;
	};

	/** True when `path` only differs from `base` by fields that never reach a consumer. */
	const isNoiseManifestChange = (path, base) => {
		if (!base) return false;
		try {
			const strip = (text) => {
				const data = JSON.parse(text);
				for (const field of IGNORED_MANIFEST_FIELDS) delete data[field];
				for (const field of DEP_FIELDS) {
					for (const dep of Object.keys(data[field] ?? {})) {
						if (names.has(dep)) delete data[field][dep];
					}
				}
				return canonical(data);
			};
			return strip(git(['show', `${base}:${path}`])) === strip(git(['show', `HEAD:${path}`]));
		} catch {
			return false;
		}
	};

	const isReleaseArtifact = (path, base) =>
		/(?:^|\/)CHANGELOG\.md$/u.test(path) ||
		(/(?:^|\/)package\.json$/u.test(path) && isNoiseManifestChange(path, base));

	const changedFiles = (base) =>
		(base ? git(['diff', '--name-only', `${base}..HEAD`]) : git(['ls-files']))
			.split('\n')
			.map((f) => f.trim())
			.filter((f) => f && isPublishedFile(f) && !isReleaseArtifact(f, base));

	/** Highest bump level among commits since `base` that touch published files in `scope`. */
	const bumpLevel = (base, scope) => {
		if (!base) return 'patch';
		const raw = git(['log', '--format=%H%x1f%s%x1f%b%x1e', `${base}..HEAD`, '--', ...scope]);
		let best = 'patch';
		for (const record of raw.split('\x1e')) {
			const [hash, subject = '', body = ''] = record.trim().split('\x1f');
			if (!hash) continue;
			const level = commitLevel(subject, body);
			if (BUMP_RANK[level] <= BUMP_RANK[best]) continue;
			const files = git(['show', hash, '--name-only', '--format='])
				.split('\n')
				.filter((f) => f && isPublishedFile(f) && touches([f], scope));
			if (files.length > 0) best = level;
			if (best === 'major') break;
		}
		return best;
	};

	const plan = {};
	for (const key of order) {
		const meta = table[key];
		const base = baselineTag(meta.npm);
		const files = changedFiles(base);
		const deps = internalDeps(meta);
		const scopeKeys = new Set([key]);
		for (const dep of deps) {
			scopeKeys.add(dep);
			plan[dep].scopeKeys.forEach((k) => scopeKeys.add(k));
		}
		const scope = [...[...scopeKeys].flatMap((k) => fullScopeOf(table[k])), ...globalTriggers];
		const via = (cond, why) => (cond ? why : null);
		const reason =
			via(!base, 'no previous tag') ||
			via(touches(files, scopeOf(meta)), 'own files changed') ||
			via(touches(files, triggersOf(meta)), 'bundled internal package changed') ||
			via(
				deps.some((d) => plan[d].release),
				'internal dependency released',
			) ||
			via(touches(files, globalTriggers), 'shared build pipeline changed');
		const release = Boolean(reason);

		const manifestVersion = readJson(manifestPath(meta)).version || '0.0.0';
		const published = npm ? npm(meta.npm) : null;
		const tagged = maxSemver(
			git(['tag', '--list', `${meta.npm}@*`])
				.split('\n')
				.map((t) => t.trim().slice(meta.npm.length + 1))
				.filter((v) => SEMVER.test(v)),
		);
		const current = maxSemver([tagged, published ?? '0.0.0', manifestVersion]);
		// Never published and never tagged: the manifest version IS the first release.
		const initial = release && !base && published === null && tagged === '0.0.0';
		const bump = !release ? null : initial ? 'initial' : bumpLevel(base, scope);
		const version = !release ? current : initial ? manifestVersion : bumpVersion(current, bump);
		plan[key] = {
			npm: meta.npm,
			dir: meta.dir,
			manifest: `${meta.dir}/package.json`,
			changelog: `${meta.dir}/CHANGELOG.md`,
			baseline: base,
			release,
			reason: release ? reason : null,
			initial,
			bump,
			currentVersion: current,
			version,
			tag: `${meta.npm}@${version}`,
			dependsOn: deps,
			scopeKeys: [...scopeKeys],
			includePaths: [
				...[...scopeKeys].flatMap((k) => [
					...(table[k].paths
						? table[k].paths.map((p) => (p.endsWith('/') ? `${p}**` : p))
						: [`${table[k].dir}/**`]),
					...triggersOf(table[k]).map((dir) => `${dir}/**`),
				]),
				...globalTriggers,
			],
		};
	}
	return { anyChanged: Object.values(plan).some((p) => p.release), order, packages: plan };
}

/**
 * Stamp each released package's version into its package.json, and repoint every internal
 * dependency range at the version being released (keeping a `^`/`~` prefix if the range has one),
 * so manifests, the lockfile and the published tarballs stay mutually installable.
 */
export function applyPlan({ root, packages: table }, plan) {
	for (const [key, meta] of Object.entries(table)) {
		const path = join(root, meta.dir, 'package.json');
		const data = readJson(path);
		let changed = false;
		if (plan.packages[key].release && data.version !== plan.packages[key].version) {
			data.version = plan.packages[key].version;
			changed = true;
		}
		for (const field of DEP_FIELDS) {
			for (const dep of Object.keys(data[field] ?? {})) {
				const target = Object.values(plan.packages).find((p) => p.npm === dep);
				if (!target?.release) continue;
				const prefix = /^[\^~]/u.exec(data[field][dep])?.[0] ?? '';
				if (data[field][dep] !== `${prefix}${target.version}`) {
					data[field][dep] = `${prefix}${target.version}`;
					changed = true;
				}
			}
		}
		if (changed) writeJson(path, data);
	}
}

function main() {
	const argv = process.argv.slice(2);
	const root = join(dirname(fileURLToPath(import.meta.url)), '..');
	const config = { root, packages: PACKAGES, globalTriggers: GLOBAL_TRIGGERS };
	const plan = planRelease({ ...config, npm: argv.includes('--no-npm') ? undefined : npmVersion });
	writeJson(join(root, 'release-plan.json'), plan);
	if (argv.includes('--write')) applyPlan(config, plan);

	console.log('Release plan (independent per-package versions):');
	for (const key of plan.order) {
		const p = plan.packages[key];
		const line = p.release
			? `${p.currentVersion} -> ${p.version} (${p.bump}; ${p.reason})  tag ${p.tag}`
			: `${p.currentVersion} (skip)`;
		console.log(`  ${key.padEnd(14)} ${line}`);
	}
	if (argv.includes('--write')) console.log('(wrote versions to released package.json files)');
	if (process.env.GITHUB_OUTPUT) {
		appendFileSync(process.env.GITHUB_OUTPUT, `any_changed=${plan.anyChanged}\n`);
	}
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
