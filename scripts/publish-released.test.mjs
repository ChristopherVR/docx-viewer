import assert from 'node:assert/strict';
import { test } from 'node:test';

import { PACKAGES } from './release-plan.mjs';
import { resolveTargets, verifyManifest } from './publish-released.mjs';

const [key, meta] = Object.entries(PACKAGES)[0];
const version = JSON.parse(
	await (await import('node:fs/promises')).readFile(`${meta.dir}/package.json`, 'utf8'),
).version;

test('a tag resolves to exactly one package, scoped names included', () => {
	const [target] = resolveTargets({ tag: `${meta.npm}@1.2.3` });
	assert.deepEqual(target, { key, npm: meta.npm, dir: meta.dir, version: '1.2.3' });
});

test('a malformed or unknown tag is rejected', () => {
	assert.throws(() => resolveTargets({ tag: 'v1.2.3' }), /Invalid tag|Unknown package/);
	assert.throws(() => resolveTargets({ tag: 'nope@1.0.0' }), /Unknown package/);
	assert.throws(() => resolveTargets({ tag: `${meta.npm}@1.0.0; rm -rf /` }), /Invalid tag/);
});

test('a plan resolves to its released packages in plan order', () => {
	const plan = {
		order: ['b', 'a', 'c'],
		packages: {
			a: { release: true, npm: '@x/a', dir: 'packages/a', version: '1.0.0' },
			b: { release: false, npm: '@x/b', dir: 'packages/b', version: '1.0.0' },
			c: { release: true, npm: '@x/c', dir: 'packages/c', version: '2.0.0' },
		},
	};
	assert.deepEqual(
		resolveTargets({ plan }).map((t) => t.npm),
		['@x/a', '@x/c'],
	);
});

test('the manifest on disk must be the version being published', () => {
	assert.throws(
		() => verifyManifest({ npm: meta.npm, dir: meta.dir, version: '9.9.9' }),
		/on disk/,
	);
	assert.doesNotThrow(() => verifyManifest({ npm: meta.npm, dir: meta.dir, version }));
});

test('exactly the packages in the release table are public; every other workspace package is private', async () => {
	const { readdir, readFile } = await import('node:fs/promises');
	const published = new Set(Object.values(PACKAGES).map((p) => p.dir.replace('packages/', '')));
	assert.equal(published.size, 7);
	for (const dir of await readdir('packages')) {
		const manifest = JSON.parse(await readFile(`packages/${dir}/package.json`, 'utf8'));
		assert.equal(
			Boolean(manifest.private),
			!published.has(dir),
			`packages/${dir}: ${published.has(dir) ? 'must be public' : 'must be private'}`,
		);
		if (published.has(dir))
			assert.doesNotThrow(() =>
				verifyManifest({ npm: manifest.name, dir: `packages/${dir}`, version: manifest.version }),
			);
	}
});
