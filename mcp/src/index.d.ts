import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
export interface ServerOptions {
	rootDir?: string;
}
export function registerTools(server: McpServer, options?: ServerOptions): void;
export function createServer(options?: ServerOptions): McpServer;
