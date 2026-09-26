# PowerPoint repository reuse audit

CFB/OLE2 primitives and Word 97-2003 binary helpers live in `D:/Development/ole2` as package `@christophervr/ole2`. The extracted source records its provenance and Apache-2.0 licensing. DOCX is a separate OOXML format and its model, parser, and serializer are canonical in Word's `@christophervr/docx-core` package. No modern DOCX code is part of `ole2`.

## Shared package API

The shared package exposes generic CFB APIs and Word binary-format modules through separate subpaths. The DOC legacy adapter uses these shared primitives and owns conversion into the document model, warning text, and the supported editing contract. It does not contain DOCX parsing or serialization.

| Candidate                               | Assessment                                                                                                                                                                                                                                                                           |
| --------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| OLE2 reader/types and writer helpers    | Generic named-stream and CFB directory/sector primitives shared by the Word and PowerPoint viewers. The package builds JS and declaration outputs and tests mini/regular streams, root CLSID, directory ordering, and DIFAT behavior. Further bounds-hardening is still appropriate. |
| Windows-1252, piece table, FIB, FKP     | Shared package placement avoids duplicate implementations, but these APIs remain Word-specific. Keep them behind the `ole-document-doc-*` subpaths rather than treating them as generic CFB APIs.                                                                                    |
| Word CFB bridge and paragraph editor    | The bridge binds CFB primitives to Word stream names and the editor understands Word character positions, FKP tables, and unsafe PLCFs. They belong in the shared binary-format package for reuse, but remain format-specific and are not a generic viewer/editing layer.            |
| Excel BIFF8 reader and cell writers     | Canonical implementation is in `ole2`; PowerPoint keeps compatibility reexports. The neutral grid model has no viewer dependency.                                                                                                                                                    |
| PPT record parsing and constants        | Canonical implementation is in `ole2`; PowerPoint imports the same record traversal and constants.                                                                                                                                                                                   |
| PPT binary export                       | Move the `WDeck` binary model and serializer dependency graph into `ole2`. Conversion from `PptxElement`, modern OOXML packaging and rendering stay in the PowerPoint consumer.                                                                                                      |
| Visio/Publisher and document properties | Binary inspection, standard OLE metadata, and preservation-aware stream edits belong in `ole2`. Application-specific page and drawing editors are separate features.                                                                                                                 |

Both repositories consume the same concrete legacy-format implementation. Compatibility modules may reexport public names but must not retain a second implementation.

The DOCX legacy API extracts main-body paragraph text, preserves exact source bytes for an unchanged save, and permits paragraph text replacement only when the editor's safety gate accepts the document. It does not promise Word layout parity or lossless editing. Encryption, malformed containers, tables, paragraph insertion/deletion, formatting edits, and documents with unsafe CP-indexed structures are outside the supported editing subset.

## DOCX consolidation

The ordered DOCX model/parser/writer and embedded-DOCX API live in Word's `@christophervr/docx-core` package. PowerPoint's current embedded-DOCX adapter remains in place; migration to the core API is deferred until the package is published and consumer integration is ready. `@christophervr/ole2` remains limited to CFB and legacy binary formats.

## Next shared package boundary

The PowerPoint checkout contains substantial reusable modern-Office infrastructure. These source counts exclude tests and include comments/generated data; they measure candidates, not code already migrated.

| Candidate                                                       |           Current size | Extraction boundary                                                                                                                                                 |
| --------------------------------------------------------------- | ---------------------: | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Colour primitives, transforms and theme-colour references       |  6 files / 1,152 lines | Separate scalar colour math from `XmlObject` and `PptxThemeColorRef` adapters. Word must use the same transform semantics when adding theme colours.                |
| Font helpers and advance-width data                             | 14 files / 2,385 lines | Share measured glyph advances and font metadata once Word implements pagination/text measurement; preserve source-data provenance and regenerate from one pipeline. |
| Framework-neutral theme types, presets and CSS variable mapping |    7 files / 607 lines | Parameterize the CSS token prefix and retain `--pptx-*` compatibility. One theme package can then feed both web components and documentation.                       |

A separate `office-shared` package is the appropriate future home for those modern-Office utilities; placing UI themes or OOXML semantics in `ole2` would obscure the legacy-format boundary. Extract each feature with a real consumer in both viewers, move its tests, and replace the original implementation with a compatibility adapter. Do not copy the large PowerPoint runtime or import its document model into Word merely to reuse a helper.

Word's editing commands and model-to-DOM conversion live in its single web component. React, Vue, Angular, Svelte and vanilla bindings remain lifecycle/event adapters. Tailwind is part of the demo build, while the component carries its own scoped CSS so host frameworks do not need a Tailwind build to render it.
