#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './index.js';

const server = createServer({ rootDir: process.argv[2] ?? process.cwd() });
await server.connect(new StdioServerTransport());
