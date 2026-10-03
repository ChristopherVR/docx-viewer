# docx-viewer-mcp

Repository-owned MCP schemas and server wiring. All document loading, editing,
calculation and serialization delegate to `ooxml-core/automation`.
Requires the core release that introduces that entry point (0.11.0 or later).

```json
{
	"mcpServers": {
		"docx": {
			"command": "npx",
			"args": ["-y", "--package=docx-viewer-mcp", "docx-tools", "/path/to/documents"]
		}
	}
}
```

Paths are scoped to the configured root (the current directory by default).
Creation and save-as never overwrite existing files. Edit tools update the source
unless `outputPath` names a new file. Files are written atomically after core
validation. Unsupported features and preservation limitations are returned by core.

Programmatic integration: `createServer({ rootDir })` or
`registerTools(existingMcpServer, { rootDir })`. Importing the package does not
start a transport. The OOXML combined MCP calls this same registration function.

This package ships JavaScript source and needs no separate build step.
