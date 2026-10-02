import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
	forbiddenManifestEntries,
	importedPackages,
	undeclaredImports,
} from './check-published-refs.mjs';

const manifest = {
	dependencies: { 'docx-core': '^0.1.0', 'prosemirror-state': '^1.4.3' },
	peerDependencies: { react: '>=18' },
};

test('finds static, dynamic and re-export specifiers but skips relative ones', () => {
	const source = [
		"import a from 'react';",
		'export * from "prosemirror-state";',
		"const b = await import('@christophervr/ole2/ole2-parser-read');",
		"import './chunk.js';",
		'const msg = \'x from": "y\';',
	].join('\n');
	assert.deepEqual(importedPackages(source), [
		'react',
		'prosemirror-state',
		'@christophervr/ole2/ole2-parser-read',
	]);
});

test('declared dependencies, subpaths and node built-ins are allowed', () => {
	const source = [
		"import { x } from 'docx-core';",
		"import 'prosemirror-state/dist/index.js';",
		"import fs from 'node:fs';",
		"import 'react';",
	].join('\n');
	assert.deepEqual(undeclaredImports(source, manifest), []);
});

test('flags inlined internal packages and ole2 that leaked into a tarball', () => {
	const source = [
		"import 'prosemirror-view';",
		"export * from 'docx-web-component';",
		"import '@christophervr/ole2/ole-document-doc-fib';",
	].join('\n');
	assert.deepEqual(undeclaredImports(source, manifest), [
		'prosemirror-view',
		'docx-web-component',
		'@christophervr/ole2/ole-document-doc-fib',
	]);
});

test('rejects internal, ole2 and workspace entries in a published manifest', () => {
	assert.deepEqual(forbiddenManifestEntries(manifest), []);
	// ooxml-ui is a registry dependency like docx-core: never inlined, never forbidden.
	assert.deepEqual(
		forbiddenManifestEntries({ dependencies: { 'ooxml-ui': '^0.1.1' } }),
		[],
	);
	assert.deepEqual(
		forbiddenManifestEntries({
			dependencies: {
				'@christophervr/ole2': '0.2.0',
				'docx-legacy': '^0.1.0',
				'docx-core': 'workspace:*',
			},
		}),
		['dependencies.@christophervr/ole2', 'dependencies.docx-legacy', 'dependencies.docx-core'],
	);
});
