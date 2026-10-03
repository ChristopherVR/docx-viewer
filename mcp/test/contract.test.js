import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../src/index.js';

test('docx MCP discovery, input validation and tool errors', async () => {
	const root = await mkdtemp(join(tmpdir(), 'docx-mcp-'));
	const server = createServer({ rootDir: root });
	const client = new Client({ name: 'test', version: '1.0.0' });
	const [a, b] = InMemoryTransport.createLinkedPair();
	try {
		await server.connect(b);
		await client.connect(a);
		const { tools } = await client.listTools();
		assert.ok(tools.length >= 2);
		assert.ok(tools.every((tool) => tool.name.startsWith('docx_')));
		const inspect = tools.find((tool) => tool.name === 'docx_inspect');
		assert.equal(inspect.annotations.readOnlyHint, true);
		const missing = await client.callTool({
			name: 'docx_inspect',
			arguments: { filePath: 'missing.docx' },
		});
		assert.equal(missing.isError, true);
		const invalid = await client.callTool({ name: 'docx_inspect', arguments: { filePath: 123 } });
		assert.equal(invalid.isError, true);
		const created = await client.callTool({
			name: 'docx_create',
			arguments: { filePath: 'new.docx' },
		});
		assert.ok(!created.isError);
		const duplicate = await client.callTool({
			name: 'docx_create',
			arguments: { filePath: 'new.docx' },
		});
		assert.equal(duplicate.isError, true);
		const inspected = await client.callTool({
			name: 'docx_inspect',
			arguments: { filePath: 'new.docx' },
		});
		assert.ok(!inspected.isError);
	} finally {
		await client.close();
		await server.close();
		await rm(root, { recursive: true, force: true });
	}
});
