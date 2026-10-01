# Architecture and ownership

## Dependency direction

```text
React | Vue | Angular | Svelte | vanilla bindings
                         |
                  web component editor
                         |
                 document detection
                  /              \
       @christophervr/docx-core   legacy DOC adapter
                                       |
                           @christophervr/ole2
                            CFB and binary DOC only
```

The canonical document model, DOCX parser, and serializer live in the `docx` area of the private `@christophervr/ooxml-core`; `@christophervr/docx-core` is a thin published entry point that re-exports it. The editable model is a deliberately small projection of OOXML. A loaded session owns the original package plus the baseline model, so saving supported edits can preserve parts outside that projection. No-op save returns original bytes.

`document` detects container signatures and selects DOCX or legacy DOC. `legacy` extracts main-body paragraphs from Word 97-2003 files, uses `ole2` for CFB and binary Word structures, and translates extracted text into the common model. Its editing path accepts plain paragraph text only when the safety checks pass.

`web-component` is the only editing UI. It converts the model into ProseMirror, applies editing transactions and repairs block identities, emits model changes, and calls the loaded session's save function. The current command surface supports paragraph editing, basic direct text formatting, alignment, simple tables, and zoom on a continuous editing surface. It does not implement Word pagination. Shadow DOM scopes the toolbar and page styles. Registration is browser-only and idempotent. Browser tests exercise the same contract through every binding.

`bindings` owns framework integration only: mount/unmount, property updates, event forwarding and imperative refs. New editor features belong in the web component; a wrapper change is justified only by a framework lifecycle/API requirement.

## Sharing across the two viewers

CFB reader/writer and Word binary internals are canonical in the sibling `../ole2` package `@christophervr/ole2`. DOCX model, parser, writer, and the planned embedded-DOCX API are canonical in `@christophervr/ooxml-core/docx` (re-exported by `@christophervr/docx-core`). PowerPoint keeps its existing embedded-DOCX adapter until the Word core is published and its consumer migration is ready. No modern DOCX implementation belongs in `ole2`.

The shared `ole2` package has no React, Word, or PowerPoint UI dependency and owns only CFB and legacy Word binary formats. DOCX is OOXML and belongs to `ooxml-core/docx`.

## Build and release boundary

The current release scope is `@christophervr/ole2` only. Word package publication and PowerPoint's migration to `@christophervr/docx-core/embedded` are deferred until the shared dependency and consumer integration are ready. The shared CFB package contains no DOCX codec.

Bundled dependencies are included in application package output where needed, so downstream consumers do not need sibling source checkouts.

The package names in the integration guides use the `@christophervr` registry scope. The Pages examples are built from this workspace and do not imply the Word packages have been published.

## Validation recorded during bootstrap

- Shared `ole2`: build and 27 tests passed for CFB and Word binary support.
- Word core: DOCX model, parser, serializer, and preservation regressions passed in the package suite.
- Word workspace: typecheck, formatting, production demo build and 21 tests passed.
- Browser contract: all five framework mounts plus legacy DOC import/edit/save/reopen/export passed (6 tests, Chromium).
- PowerPoint: core typecheck/build, 132 initial binary/OLE/crypto regressions and 17 DOCX/embedded-operation regressions passed (the two targeted runs overlap). Built ESM import and a scan for leaked shared-package/XML-DOM runtime imports passed.
- PowerPoint's entire test suite, all five PowerPoint UI builds, other browser engines and real Word visual pagination comparisons were not run.

Review the parity roadmap before treating this as Word feature parity. Successful package builds and demo routes do not establish Word rendering, pagination, or round-trip parity.
