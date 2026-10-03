import { inspectDocx, createDocx, setDocxRunText } from 'ooxml-core/automation';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { createFileOperations } from 'ooxml-core/automation/node';

export function registerTools(server, options = {}) {
	const files = createFileOperations(options.rootDir);
	const path = z.string().min(1);
	const fileSchema = { filePath: path };
	const editSchema = { ...fileSchema, outputPath: path.optional() };
	const register = (name, description, inputSchema, readOnlyHint, handler) =>
		server.registerTool(
			name,
			{
				description,
				inputSchema,
				annotations: { readOnlyHint, destructiveHint: !readOnlyHint, openWorldHint: false },
			},
			async (params) => {
				try {
					const result = await handler(params);
					return { content: [{ type: 'text', text: JSON.stringify(result) }] };
				} catch (error) {
					return {
						isError: true,
						content: [
							{ type: 'text', text: error instanceof Error ? error.message : String(error) },
						],
					};
				}
			},
		);

	register(
		'docx_inspect',
		'Inspect Word blocks, stable paragraph ids, runs and warnings.',
		fileSchema,
		true,
		(p) => files.inspect(p.filePath, ['.docx'], inspectDocx),
	);
	register(
		'docx_create',
		'Create a Word document at a new path. Existing files are never overwritten.',
		{ ...fileSchema, paragraphs: z.array(z.string()).max(10000).default([]) },
		false,
		(p) => files.create(p.filePath, ['.docx'], () => createDocx(p.paragraphs)),
	);
	register(
		'docx_set_run_text',
		'Set one ordinary text run, preserving its formatting. Fields, images, equations and revisions are rejected. Without outputPath the source is updated.',
		{
			...editSchema,
			paragraphId: path,
			runIndex: z.number().int().min(0),
			text: z.string().max(1000000),
		},
		false,
		(p) =>
			files.edit(
				p.filePath,
				['.docx'],
				(bytes) => setDocxRunText(bytes, p.paragraphId, p.runIndex, p.text),
				p.outputPath,
			),
	);
}

export function createServer(options = {}) {
	const server = new McpServer({ name: 'docx-viewer-tools', version: '0.1.0' });
	registerTools(server, options);
	return server;
}
