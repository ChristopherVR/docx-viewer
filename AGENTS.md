# Working agreements

- One framework-neutral document model and one web-component editor. UI bindings are lifecycle/event adapters only.
- Keep source modules below 300 lines where practical. Add meaningful regression tests for parsing, preservation, editing and binding contracts.
- Unsupported document features must be reported honestly. Never silently claim Word layout parity or lossless export.
- Shared OLE and legacy Word implementations live in ../ole2. Both viewers consume that package; do not fork these implementations. PowerPoint may be changed to integrate shared code while preserving its public API. Record extraction provenance.
- Use Bun, TypeScript strict mode, Vitest and browser contract tests. No commits or publishing unless requested.

- DOCX model, parser and serializer live in packages/core. Its /embedded entry is ready for a future PowerPoint migration after Word packages are published. Keep ole2 restricted to legacy compound-file codecs; never move modern OOXML handling into it.
- Shared modern OOXML logic (units now; colour, geometry, packaging, XML, DrawingML, charts, diagrams later) lives in the separate private repository `../ooxml-core` (a sibling of this repo, like ole2). Consume its packages (`@christophervr/ooxml-*`); do not copy or fork them. The extraction plan is `docs/ooxml-core-plan.md`. Development resolves the package from `../ooxml-core/packages/<name>` through a `file:` dependency, so build ooxml-core (`bun run build`) and re-run `bun install --force` after changing it. CI clones it at the revision pinned in `.github/workflows` (`OOXML_CORE_REF`) using the `OOXML_CORE_TOKEN` secret.

