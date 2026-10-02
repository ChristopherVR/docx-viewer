# Architecture and ownership

## Dependency direction

```text
published, one per framework (self-contained bundles)
  docx-react-viewer | -vue-viewer | -angular-viewer | -svelte-viewer | -solid-viewer | -vanilla-viewer
  ------------------------------------------------------------------------------------------------
  internal, private, inlined into every bundle:
        bindings (lifecycle and events only)
                         |
                  web-component editor (canvas text measurer)
                         |
        @christophervr/ooxml-core (external, published)
   /docx/layout   /docx/load (format detection, legacy .doc;
   /collab        ole2 codecs inlined by the core)
        ^
   docx-core (external): thin re-export of ooxml-core/docx and /docx/embedded
```

Only `@christophervr/docx-core` and the six `*-viewer` packages are published. `web-component` and `bindings` are `private` workspace packages: they are the shared source of the editor, never an npm install target. `scripts/build-packages.mjs` bundles them into each framework package, so a tarball imports only `@christophervr/docx-core`, `@christophervr/ooxml-core`, the ProseMirror libraries and its framework peers. `scripts/check-published-refs.mjs` and the pack smoke test fail the build if an internal package or `ole2` leaks into a tarball.

The canonical document model, DOCX parser, and serializer live in the `docx` area of the private `@christophervr/ooxml-core`; `@christophervr/docx-core` is a thin published entry point that re-exports it. The editable model is a deliberately small projection of OOXML. A loaded session owns the original package plus the baseline model, so saving supported edits can preserve parts outside that projection. No-op save returns original bytes.

`@christophervr/ooxml-core/docx/load` detects container signatures and selects DOCX or legacy DOC. Its legacy reader extracts main-body paragraphs from Word 97-2003 files, uses the ole2 codecs (inlined into the core) for CFB and binary Word structures, and translates extracted text into the common model. Its editing path accepts plain paragraph text only when the safety checks pass. `@christophervr/ooxml-core/docx/layout` is the DOM-free pagination engine; the editor injects its canvas-backed `TextMeasurer`. Real-time collaboration helpers (validation, identity, ordering rules) come from `@christophervr/ooxml-core/collab`; the `prosemirror-collab` client and authority, the presence plugin and the decorations stay in `web-component`.

`web-component` is the only editing UI. It converts the model into ProseMirror, applies editing transactions and repairs block identities, emits model changes, and calls the loaded session's save function. The current command surface supports paragraph editing, basic direct text formatting, alignment, simple tables, and zoom on a continuous editing surface. It does not implement Word pagination. Shadow DOM scopes the toolbar and page styles. Registration is browser-only and idempotent. Browser tests exercise the same contract through every binding.

`bindings` owns framework integration only: mount/unmount, property updates, event forwarding and imperative refs. New editor features belong in the web component; a wrapper change is justified only by a framework lifecycle/API requirement.

## Sharing across the two viewers

CFB reader/writer and Word binary internals are canonical in the sibling `../ole2` package `@christophervr/ole2`. DOCX model, parser, writer, and the planned embedded-DOCX API are canonical in `@christophervr/ooxml-core/docx` (re-exported by `@christophervr/docx-core`). PowerPoint keeps its existing embedded-DOCX adapter until the Word core is published and its consumer migration is ready. No modern DOCX implementation belongs in `ole2`.

The shared `ole2` package has no React, Word, or PowerPoint UI dependency and owns only CFB and legacy Word binary formats. DOCX is OOXML and belongs to `ooxml-core/docx`.

## Build and release boundary

Seven packages are published (`docx-core` and the six framework packages; see [releasing](/releasing)), and nothing has been released yet. PowerPoint's migration to `@christophervr/docx-core/embedded` is deferred until the Word packages are published and its consumer integration is ready. The shared CFB package contains no DOCX codec.

`@christophervr/docx-core` stays a real dependency of every framework package rather than being bundled: applications import `createDocument` and the model types from it directly, so one shared copy keeps `DocumentModel` identical on both sides. The ProseMirror libraries are real dependencies too, not bundled, so a package manager can dedupe one copy for the editor and for any other ProseMirror code in the application. `@christophervr/ooxml-core` is a real dependency too (it brings layout, loading and collab). Everything else internal is bundled so downstream consumers need no sibling source checkouts and no unpublished packages; ole2 is inlined inside the core and is not a dependency of anything here.

The package names in the integration guides use the `@christophervr` registry scope. The Pages examples are built from this workspace and do not imply the Word packages have been published.

## Validation recorded during bootstrap

- Shared `ole2`: build and 27 tests passed for CFB and Word binary support.
- Word core: DOCX model, parser, serializer, and preservation regressions passed in the package suite.
- Word workspace: typecheck, formatting, production demo build and 21 tests passed.
- Browser contract: all five framework mounts plus legacy DOC import/edit/save/reopen/export passed (6 tests, Chromium).
- PowerPoint: core typecheck/build, 132 initial binary/OLE/crypto regressions and 17 DOCX/embedded-operation regressions passed (the two targeted runs overlap). Built ESM import and a scan for leaked shared-package/XML-DOM runtime imports passed.
- PowerPoint's entire test suite, all five PowerPoint UI builds, other browser engines and real Word visual pagination comparisons were not run.

Review the parity roadmap before treating this as Word feature parity. Successful package builds and demo routes do not establish Word rendering, pagination, or round-trip parity.
