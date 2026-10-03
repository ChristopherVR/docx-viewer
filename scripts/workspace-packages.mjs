import { readdir } from 'node:fs/promises';
import assert from 'node:assert/strict';

/** Old workspaces can leave ignored build/install artifacts after their manifests are removed. */
export async function workspacePackageNames(directory) {
	const names = [];
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		assert(entry.isDirectory(), `Unexpected file in packages/: ${entry.name}`);
		const entries = await readdir(new URL(`${entry.name}/`, directory));
		if (entries.includes('package.json')) {
			names.push(entry.name);
			continue;
		}
		assert(
			entries.every((name) => name === 'dist' || name === 'node_modules'),
			`packages/${entry.name} contains source or configuration without a package.json`,
		);
	}
	return names.sort();
}
