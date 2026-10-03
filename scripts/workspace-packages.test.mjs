import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, sep } from 'node:path';
import { pathToFileURL } from 'node:url';
import { workspacePackageNames } from './workspace-packages.mjs';

async function fixture(t) {
	const path = await mkdtemp(join(tmpdir(), 'docx-workspace-packages-'));
	t.after(() => rm(path, { recursive: true, force: true }));
	return { path, url: pathToFileURL(path + sep) };
}

test('classifies manifest-bearing packages and ignores abandoned build/install artifacts', async (t) => {
	const { path, url } = await fixture(t);
	await mkdir(join(path, 'core'));
	await writeFile(join(path, 'core', 'package.json'), '{}');
	await mkdir(join(path, 'retired', 'dist'), { recursive: true });
	await mkdir(join(path, 'retired', 'node_modules'), { recursive: true });
	await writeFile(join(path, 'retired', 'dist', 'index.js'), 'build output');
	assert.deepEqual(await workspacePackageNames(url), ['core']);
});

test('still reports a new package so the classification guard rejects it', async (t) => {
	const { path, url } = await fixture(t);
	await mkdir(join(path, 'accidental'));
	await writeFile(join(path, 'accidental', 'package.json'), '{}');
	assert.deepEqual(await workspacePackageNames(url), ['accidental']);
});

test('rejects source-bearing directories that are missing their manifest', async (t) => {
	const { path, url } = await fixture(t);
	await mkdir(join(path, 'forked', 'src'), { recursive: true });
	await writeFile(join(path, 'forked', 'src', 'index.ts'), 'export const fork = true;');
	await assert.rejects(
		workspacePackageNames(url),
		/contains source or configuration without a package.json/,
	);
});

test('rejects unexpected package-root files', async (t) => {
	const { path, url } = await fixture(t);
	await writeFile(join(path, 'copied-logic.ts'), 'export const fork = true;');
	await assert.rejects(workspacePackageNames(url), /Unexpected file in packages/);
});
