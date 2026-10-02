import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { after, before, test } from 'node:test';

import {
	applyPlan,
	bumpVersion,
	commitLevel,
	isPublishedFile,
	planRelease,
} from './release-plan.mjs';

// A throwaway monorepo: core <- mid <- top, plus an independent `solo`.
const table = {
	core: { dir: 'packages/core', npm: '@x/core' },
	mid: { dir: 'packages/mid', npm: '@x/mid' },
	top: { dir: 'packages/top', npm: '@x/top' },
	solo: { dir: 'packages/solo', npm: '@x/solo' },
};
const deps = { mid: { '@x/core': '1.0.0' }, top: { '@x/mid': '1.0.0', '@x/core': '1.0.0' } };
let root;
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8' }).trim();
const write = (file, text) => {
	mkdirSync(dirname(join(root, file)), { recursive: true });
	writeFileSync(join(root, file), text);
};
const commit = (message, files) => {
	for (const [file, text] of Object.entries(files)) write(file, text);
	git('add', '-A');
	git('commit', '-q', '-m', message);
};
const plan = (npm) => planRelease({ root, packages: table, npm });
const released = (p) => p.order.filter((k) => p.packages[k].release);
let counter = 0;
const touch = (dir) => ({ [`${dir}/src/index.ts`]: `export const n = ${counter++};\n` });

before(() => {
	root = mkdtempSync(join(tmpdir(), 'release-plan-'));
	git('init', '-q', '-b', 'main');
	git('config', 'user.email', 't@example.com');
	git('config', 'user.name', 'T');
	git('config', 'commit.gpgsign', 'false');
	git('config', 'tag.gpgsign', 'false');
	const files = {};
	for (const [key, meta] of Object.entries(table)) {
		files[`${meta.dir}/package.json`] =
			`${JSON.stringify({ name: meta.npm, version: '1.0.0', dependencies: deps[key] ?? {} }, null, '\t')}\n`;
		Object.assign(files, touch(meta.dir));
	}
	commit('feat: initial', files);
});
after(() => rmSync(root, { recursive: true, force: true }));

const tagAll = () => {
	for (const meta of Object.values(table)) git('tag', '-f', `${meta.npm}@1.0.0`);
};

test('an unpublished, untagged package releases at its manifest version', () => {
	const p = plan(() => null);
	assert.deepEqual(released(p), ['core', 'mid', 'top', 'solo']);
	assert.equal(p.packages.core.initial, true);
	assert.equal(p.packages.core.version, '1.0.0');
	assert.equal(p.packages.core.tag, '@x/core@1.0.0');
});

test('a published package without a tag bumps from the registry version', () => {
	const p = plan((name) => (name === '@x/solo' ? '1.4.0' : null));
	assert.equal(p.packages.solo.initial, false);
	assert.equal(p.packages.solo.version, '1.4.1');
});

test('packages come out in dependency order', () => {
	const { order } = plan(() => null);
	assert.ok(
		order.indexOf('core') < order.indexOf('mid') && order.indexOf('mid') < order.indexOf('top'),
	);
});

test('a tagged HEAD is a no-op', () => {
	tagAll();
	const p = plan(() => '1.0.0');
	assert.equal(p.anyChanged, false);
	assert.deepEqual(released(p), []);
});

test('a test-only or changelog-only change releases nothing', () => {
	commit('feat(core): add tests', {
		'packages/core/src/index.test.ts': 'test\n',
		'packages/core/CHANGELOG.md': '# Changelog\n',
	});
	assert.equal(plan(() => '1.0.0').anyChanged, false);
});

test('a scripts-only manifest change is not a release trigger', () => {
	const path = 'packages/solo/package.json';
	const manifest = JSON.parse(readFileSync(join(root, path), 'utf8'));
	commit('chore(release): bump', {
		[path]: `${JSON.stringify({ ...manifest, scripts: { x: 'y' } }, null, '\t')}\n`,
	});
	assert.equal(plan(() => '1.0.0').anyChanged, false);
});

test('an external dependency change releases only the package whose manifest names it', () => {
	// ooxml-ui and ooxml-core are registry dependencies, not workspace packages: like a docx-core
	// range they reach consumers only through the manifest that declares them.
	const path = 'packages/solo/package.json';
	const manifest = JSON.parse(readFileSync(join(root, path), 'utf8'));
	commit('fix(solo): depend on the shared UI controls', {
		[path]: `${JSON.stringify({ ...manifest, dependencies: { '@ext/ui': '^0.1.1' } }, null, '\t')}\n`,
	});
	assert.deepEqual(released(plan(() => '1.0.0')), ['solo']);
	git('reset', '-q', '--hard', 'HEAD~1');
});

test('reordering manifest keys is not a release trigger', () => {
	const path = 'packages/solo/package.json';
	const manifest = JSON.parse(readFileSync(join(root, path), 'utf8'));
	const reordered = Object.fromEntries(Object.entries(manifest).reverse());
	commit('style(solo): sort the manifest keys', {
		[path]: `${JSON.stringify(reordered, null, '	')}
`,
	});
	assert.equal(plan(() => '1.0.0').anyChanged, false);
});

test('a fix in a leaf releases only that package, as a patch', () => {
	commit('fix(solo): handle empty input', touch('packages/solo'));
	const p = plan(() => '1.0.0');
	assert.deepEqual(released(p), ['solo']);
	assert.equal(p.packages.solo.bump, 'patch');
	assert.equal(p.packages.solo.version, '1.0.1');
	git('tag', '@x/solo@1.0.1');
});

test('a feat in a dependency re-releases every dependent with the same level', () => {
	commit('feat(core): add a parser', touch('packages/core'));
	const p = plan(() => '1.0.0');
	assert.deepEqual(released(p), ['core', 'mid', 'top']);
	for (const key of ['core', 'mid', 'top']) {
		assert.equal(p.packages[key].bump, 'minor');
		assert.equal(p.packages[key].version, '1.1.0');
	}
	assert.equal(p.packages.mid.reason, 'internal dependency released');
	assert.equal(p.packages.solo.release, false);
});

test('applyPlan stamps versions and repoints internal ranges only', () => {
	const p = plan(() => '1.0.0');
	applyPlan({ root, packages: table }, p);
	const read = (dir) => JSON.parse(readFileSync(join(root, dir, 'package.json'), 'utf8'));
	assert.equal(read('packages/core').version, '1.1.0');
	assert.deepEqual(read('packages/mid').dependencies, { '@x/core': '1.1.0' });
	assert.deepEqual(read('packages/top').dependencies, { '@x/mid': '1.1.0', '@x/core': '1.1.0' });
	assert.equal(read('packages/solo').version, '1.0.0');
});

test('the release commit itself does not retrigger a release once tagged', () => {
	git('add', '-A');
	git('commit', '-q', '-m', 'chore(release): bump versions and update changelogs [skip ci]');
	for (const key of ['core', 'mid', 'top']) git('tag', `${table[key].npm}@1.1.0`);
	assert.equal(plan(() => '1.1.0').anyChanged, false);
});

test('a breaking change bumps major, and only counts when it touches published files', () => {
	commit('feat(mid)!: drop the legacy entry', { 'packages/mid/src/index.test.ts': 'x\n' });
	assert.equal(plan(() => '1.1.0').anyChanged, false);
	commit('feat(mid)!: drop the legacy entry', touch('packages/mid'));
	const p = plan(() => '1.1.0');
	assert.deepEqual(released(p), ['mid', 'top']);
	assert.equal(p.packages.mid.version, '2.0.0');
	assert.equal(p.packages.top.version, '2.0.0');
});

test('helpers', () => {
	assert.equal(bumpVersion('1.2.3', 'minor'), '1.3.0');
	assert.equal(commitLevel('fix: x', 'BREAKING CHANGE: y'), 'major');
	assert.equal(commitLevel('feat(core): x'), 'minor');
	assert.equal(commitLevel('docs: x'), 'patch');
	assert.equal(isPublishedFile('packages/core/src/a.test.ts'), false);
	assert.equal(isPublishedFile('packages/core/src/a.ts'), true);
});

test('a change in a bundled internal directory releases the package that inlines it', () => {
	// `solo` stands in for a private internal package that `mid` bundles instead of depending on.
	const bundled = {
		core: table.core,
		mid: { ...table.mid, triggers: ['packages/solo'] },
	};
	for (const meta of Object.values(bundled)) git('tag', '-f', `${meta.npm}@9.0.0`);
	const run = () => planRelease({ root, packages: bundled, npm: () => '9.0.0' });
	assert.equal(run().anyChanged, false);
	commit('test(solo): cover empty input', { 'packages/solo/src/index.test.ts': 'x\n' });
	assert.equal(run().anyChanged, false, 'a test-only change in a trigger dir is not a release');
	commit('feat(solo): accept streams', touch('packages/solo'));
	const p = run();
	assert.deepEqual(released(p), ['mid']);
	assert.equal(p.packages.mid.reason, 'bundled internal package changed');
	assert.equal(p.packages.mid.bump, 'minor');
	assert.equal(p.packages.mid.version, '9.1.0');
	assert.ok(p.packages.mid.includePaths.includes('packages/solo/**'));
});
