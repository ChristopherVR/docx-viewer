# Working agreements

- One framework-neutral document model and one web-component editor. UI bindings are lifecycle/event adapters only.
- Keep source modules below 300 lines where practical. Add meaningful regression tests for parsing, preservation, editing and binding contracts.
- Unsupported document features must be reported honestly. Never silently claim Word layout parity or lossless export.
- Shared OLE and legacy Word implementations live in ../ole2. Both viewers consume that package; do not fork these implementations. PowerPoint may be changed to integrate shared code while preserving its public API. Record extraction provenance.
- Use Bun, TypeScript strict mode, Vitest and browser contract tests. No commits or publishing unless requested.

- DOCX model, parser and serializer live in packages/core. Its /embedded entry is ready for a future PowerPoint migration after Word packages are published. Keep ole2 restricted to legacy compound-file codecs; never move modern OOXML handling into it.
