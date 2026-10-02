# Plan: a shared `ooxml-core`

Status: proposal for review, 2026-10-01. Based on read-only reviews of `pptx-viewer-new` (main, 17d75540e, with 79 uncommitted files that were not touched), the three stale 3D worktrees, `docx-viewer/packages/core` and `ole2`. Counts are approximate and several were sampled, not exhaustively read.

## Update (2026-10-01, later): one package, all the logic

`ooxml-core` is **one published package**, `@christophervr/ooxml-core`, with an area per subpath (`/xml`, `/opc`, `/units`, `/color`, `/geometry`, and later `/drawingml`, `/chart`, `/diagram`, `/docx`, `/pptx`, `/xlsx`, `/collab`). It is the home of **all** logic for the Office products, including the document models, parsers, serializers, editing commands, layout and the Yjs collaboration and sync protocol. `docx-viewer` and `pptx-viewer` keep **only the UI**. This supersedes the multi-package table in "Proposed shape" below: read those packages as areas of the one package.

Consequences for the phases: the Word and PowerPoint cores (`docx-core`, `pptx-viewer-core`) and the collaboration logic move into the new package's `docx`, `pptx` and `collab` areas, not only the shared layers; the viewers' own packages shrink to adapters and views. The GitHub secret `OOXML_CORE_TOKEN` now exists for docx-viewer CI.

## Decisions (confirmed 2026-10-01)

1. **Repositories:** one repository per Office type (`docx-viewer`, `pptx-viewer`, `xlsx-viewer` later), plus a **separate `ooxml-core` repository**, a sibling of `ole2` at `D:\Development\ooxml-core`. It is scaffolded (strict TypeScript, Vitest, oxfmt, `AGENTS.md`, `PROVENANCE.md`) with its first package, `@christophervr/ooxml-units`. Nothing is committed or published yet.
2. **XML:** the two models are unified into one shared structure, not bridged by adapters (this replaces the adapter approach below).
3. **pptx naming:** `pptx-viewer-core` stays as a published name but becomes a thin redirect to the new core going forward.
4. **Uncommitted pptx work:** the 34 modified `.pptx` fixtures under `e2e/fixtures` were regenerated binaries and have been discarded (restored from HEAD). The remaining 45 files are an in-progress refactor of the shared Draw ribbon (UI code in `packages/shared`, `react`, `vue`, `angular`, `svelte`, `vanilla`, plus docs and a new e2e spec), not core logic; they are untouched and still need an owner decision.
5. **Goal:** a unified Office suite. Word, PowerPoint and later Excel share one OOXML core and one structure; reducing duplication is a consequence, not the aim.

## Earlier decisions

- A **new project** holds shared modern OOXML. `ole2` stays legacy-only (CFB, binary DOC/XLS/PPT, RC4/MD4). Nothing modern moves into it, and the reverse.
- Word and PowerPoint keep their own document cores (`docx-core`, `pptx-viewer-core`). They consume `ooxml-core` for what is not specific to either.
- Public APIs of both viewers are preserved (re-export shims) throughout.

## What the review found

**pptx-viewer** (`packages/core`, ~243k LOC; `packages/shared`, ~320k LOC)

- A monolith around a `PptxHandler` mixin chain whose base class owns the JSZip instance, the `fast-xml-parser` instance and theme/rels maps. Generic DrawingML code sits in the same flat folders as slide code; only file names separate them. Root `index.ts` re-exports everything, so generic helpers are already public API.
- Generic and reusable, roughly by cleanliness of the boundary:
  - **Pure, no XML library:** `color/` (1.2k), `geometry/` (22.8k preset shapes, guide formulas, custom geometry, connectors, boolean ops), `smartart-engine/` (6-8k layout algorithms), OMML to LaTeX/MathML, font metrics and substitution, media-duration, png/gif/tiff utilities, `ooxml-crypto*` (1.6k).
  - **Parse/serialize behind a part loader:** charts (~15k in `utils/chart-*`, plus chartex, colour/style parts, embedded workbook), SmartArt (~32k in `utils/smartart-*`, drawing part, fabrication), fills/lines/effects, text-body parsers, theme, VML, InkML.
  - **Rendering in `shared`:** chart view-model and 3D (~30k), SmartArt 3D (~8k), `three-view` host. These take a whole `PptxElement`; they need only a chart or diagram model.
- Slide-only (stays): slides/masters/layouts/notes, animation and transitions, presentation properties, the fluent builder SDK, exporters, bindings.
- Not shareable: `a:tbl` (Word has `w:tbl`), `a:rPr` runs (Word has `w:rPr`), numbering, borders. Only token tables and primitives overlap.
- 3D: `three-d-parity-push` is merged into main. `three-d-charts`, `three-d-smartart` and `three-d-bindings` are ~587 commits behind, conflict in the same files, and are superseded; treat as retired. Ground truth for rendering is small (17 chart slides, 112 SmartArt slides over 8 layouts).

**docx-viewer** (`packages/core`, ~120 flat modules)

- Strict TypeScript (`exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`), branded units (Twips, Emu...), `@xmldom/xmldom` 0.9 DOM, schema-ordered writes that patch only changed XML, generated `ST_*` types from ECMA-376 XSDs, xmllint-wasm schema-validity tests, ESM only.

**Duplication that is real today**

- Units/EMU constants, relationship parsing/allocation, `[Content_Types].xml` handling, core/app properties, theme colour scheme, hyperlink safe-href, number-format tables, validation issue shape, MCE (`mc:AlternateContent`) handling, OOXML agile-encryption primitives.
- `ole2` holds a copy of the modern OOXML crypto primitives and a stale comment about a `docx` editor. PPT record parsing exists in both `ole2/src/ppt` and pptx core and has drifted. docx pins `ole2` 0.2.0, pptx 0.3.0.

## The central problem: two incompatible XML models, to be unified

pptx uses `fast-xml-parser` object trees with schema reordering; docx uses DOM nodes. Because the suite needs one model, the recommendation is to unify on an **ordered, namespace-aware node tree with preservation** (the docx approach), not on object trees:

- Word mixed content (`w:p` / `w:r`) and ordered children cannot be represented faithfully by object trees, which group same-name siblings and need per-family reorder tables (pptx has `xml-reorder`, `ordered-xml-merge` and several `*-sibling-order` modules for exactly this).
- DOM-style nodes keep unknown elements, attributes, comments and order for free, which is what preservation needs.

Sequence, so nothing breaks while it happens:

1. **Extract the packages that do not touch XML first** (units, colour, geometry, maths, crypto, layout engines). Both products can use them immediately.
2. **Build `ooxml-xml`**: the unified node tree, parser, serializer, namespace and schema-order helpers, with the docx behaviour (DTD/entity rejection, strict parse errors) as the baseline. Provide a temporary **object-tree compatibility view** so pptx code that still reads `XmlObject` keeps working while it is ported area by area.
3. **Port pptx area by area** to `ooxml-xml` (charts and SmartArt first, because they are being shared; slides, text and tables later), deleting its reorder machinery as each area moves. This is the largest cost in the plan: pptx core is ~243k lines and `XmlObject` is threaded through all of it.
4. Remove the compatibility view once nothing uses it, and drop `fast-xml-parser` from the core.

## Proposed shape

One repository, `ooxml-core`, Bun workspaces, Apache-2.0, ESM + CJS (pptx publishes both). Scope `@christophervr/ooxml-*`, one package per boundary so consumers pull only what they use:

| Package                                                  | Contents                                                                                                                               | Depends on                     |
| -------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------ |
| `ooxml-units`                                            | branded Emu/Twips/points, EMU constants, rounding                                                                                      | none                           |
| `ooxml-color`                                            | colour parsing and DrawingML transforms (lum/sat/tint/shade)                                                                           | units                          |
| `ooxml-geometry`                                         | preset shapes, guide evaluator, custom geometry, connectors                                                                            | units                          |
| `ooxml-opc`                                              | zip package, parts, content types, relationships, allocator, core/app properties, MCE, strict-to-transitional map, validator (OPC/MCE) | jszip                          |
| `ooxml-xml`                                              | `XmlReader/Writer` interface, DOM adapter, object-tree adapter, schema-order helpers, entity/whitespace helpers                        | none (adapters optional peers) |
| `ooxml-drawingml`                                        | models + parse/serialize: fills, lines, effects, text body (shared subset), theme, blip/xfrm, VML, InkML, fonts/embedded fonts         | xml, opc, color, geometry      |
| `ooxml-chart`                                            | chart models, ChartML/chartex parse and serialize, colour/style parts, embedded workbook, 2D view-model                                | drawingml                      |
| `ooxml-diagram`                                          | SmartArt models, data/layout/colors/quickStyle/drawing parts, layout engine, fabrication                                               | drawingml                      |
| `ooxml-chart-3d`, `ooxml-diagram-3d`, `ooxml-three-view` | three.js scene builders and host (three stays an optional lazy peer)                                                                   | chart, diagram                 |
| `ooxml-math`                                             | OMML <-> LaTeX/MathML                                                                                                                  | none                           |
| `ooxml-crypto`                                           | agile/standard encryption primitives (moved from ole2/pptx copies)                                                                     | none                           |
| `ooxml-schema`                                           | the XSD-driven type generator, extended from WML to DML, PML, chart, diagram                                                           | build-time only                |

A part loader interface (read XML part, resolve a relationship, read binary, theme and colour-map lookup) replaces `slidePath`/`PptxHandlerRuntime` in chart and SmartArt parsing. Word implements it over `word/document.xml` rels (`c:chart r:id`, `dgm:relIds`); PowerPoint over its slide rels. Renderers take narrow `ChartModel`/`DiagramModel` inputs, not `PptxElement`.

Engineering rules carried over from this repo: modules under 300 lines where practical, strict TypeScript for new packages, regression tests for parsing/preservation, unsupported features reported honestly, extraction provenance recorded, no commits or publishing without being asked.

## Phases

**Phase 0: ground rules (days).** Create the repo and CI. Fix tsconfig policy (strict for all new packages; products may stay looser and consume `.d.ts`). Settle xmldom 0.9 vs 0.8, ESM+CJS output, and package naming. Decide how pptx's 79 uncommitted files are handled before touching that tree (they belong to someone's in-progress work).

**Phase 1: pure packages (1-2 weeks).** `ooxml-units`, `ooxml-color`, `ooxml-geometry`, `ooxml-math`, `ooxml-crypto`, font metrics. Move with their existing tests; pptx re-exports from its old paths so its API does not change; docx adopts units and colour where it has gaps. Remove the crypto copy and stale comment from `ole2`.
Exit: both products build and pass tests on the shared packages; no behaviour change.

**Phase 2: OPC + unified XML (3-4 weeks).** `ooxml-opc` from docx's relationship/allocator/content-type/core-property code (pure, tested) plus pptx's validator and MCE helpers; `ooxml-xml` as the single node tree with the temporary object-tree compatibility view. docx adopts it directly (it already uses DOM nodes); pptx starts consuming it through the compatibility view.
Exit: round-trip byte-identity tests (docx no-op saves, pptx fixtures) unchanged.

**Phase 3: DrawingML shared layer (4-6 weeks, includes the first pptx ports).** Fills, lines, effects, theme, blip/xfrm, VML, shared text-body subset, written once against `ooxml-xml`; pptx's matching parsers are ported and its reorder code removed as each lands. This is also where docx gets a real `wps` text-box/shape model instead of the inline-only text box added now, and a drawing-anchor model for floating objects.

**Phase 4: charts and SmartArt (4-6 weeks).** Move models, parse/serialize, the layout engine and fabrication. Introduce the part loader and narrow render inputs; pptx keeps its `PptxElement` wrappers over them. docx gains inline/anchored `c:chart` and `dgm` parsing, rendering and (initially) read-only round-trip.
Exit: pptx's 17 chart and 112 SmartArt ground-truth slides unchanged (acceptance gate); docx fixtures render and survive save byte-for-byte when untouched.

**Phase 5: 3D renderers (2-3 weeks).** Move chart/SmartArt 3D scene code and `three-view` once their inputs are narrow models. Keep three lazy and optional.

**Phase 6: schema generation and cleanup.** Extend the generator to DML/PML/chart/diagram simple types; consolidate duplicate PPT record parsing so it lives only in `ole2`; align `ole2` version pins; retire the three stale 3D worktrees after confirming nothing unique remains.

## Risks and how to contain them

- **XML-model divergence** (above): interface plus adapters; no forced convergence.
- **pptx scale and public API:** shims at the old import paths; extract by filename-prefix with tests moving alongside; one area per PR.
- **Mixed-content ordering:** pptx's sibling-order machinery is paragraph-shaped; Word needs the same for `w:p/w:r`. Keep it in each product until `ooxml-xml` has a proven generic form.
- **Rendering regressions:** pptx ground-truth MAE numbers (whole-slide 3-4.7, worst 8.64) are the baseline; any change that moves them is rejected.
- **Strictness mismatch:** consuming strict packages from a looser project is fine; the reverse is not, so shared code is strict from the start.
- **Concurrent work in the pptx tree:** extraction PRs must not be cut from a dirty working tree.
- **Word-side honesty:** charts stay placeholders and say so. SmartArt in Word documents is shown read-only from its cached drawing (`@christophervr/office-ui`'s `<office-ui-smartart>`); it is not recomputed or editable, and its approximations are listed in the editor.

## Status (2026-10-02)

- `@christophervr/ooxml-core@0.1.0` is published to npm from the public `ChristopherVR/ooxml-core` repository (releases publish from CI with OIDC). It holds the units, colour, geometry, XML, OPC, `docx` and `pptx` areas as one package.
- docx-viewer consumes it as a normal `^0.1.0` dependency of `@christophervr/docx-core`; the pack smoke test installs it from the registry. CI no longer clones the core.
- docx-viewer publishes only `@christophervr/docx-core` (thin re-export of `/docx` and `/docx/embedded`) and six self-contained framework packages (`docx-<framework>-viewer`). The UI packages (`document`, `layout`, `legacy`, `web-component`, `bindings`) are private and bundled into each; the `@christophervr/ole2` codecs are inlined too, so `ole2` is not a dependency of anything published. See `docs/releasing.md`.
- Remaining plan: move Word layout and collaboration logic into the core, then xlsx and Visio areas.

## Next steps

1. Commit the `ooxml-core` scaffold (needs your go-ahead) and create the GitHub repository.
2. Switch docx-viewer to consume `@christophervr/ooxml-units` (replacing `packages/core/src/units.ts`, keeping the Word-model type tests here), then the same for pptx via its constants.
3. Decide the owner of the 45 uncommitted Draw-ribbon files in `pptx-viewer-new`.
4. Start `ooxml-color` and `ooxml-geometry` (pure code, pptx source), then `ooxml-opc`.
5. Write the `ooxml-xml` design note (node model, namespace handling, preservation guarantees, compatibility view) before any pptx porting.

## Questions (answered above; kept for the record)

1. Repo name and scope: `ooxml-core` as one repo with the packages above, or a single published package with subpath entries? (I recommend multiple packages.)
2. Is converging the two XML models a goal, or is the adapter approach acceptable long term?
3. Should pptx's unscoped `pptx-viewer-core` name be kept, with `ooxml-*` as new dependencies?
4. Who owns the uncommitted pptx work, and may extraction branches start from main once it is committed or stashed?
5. Priority: is Word chart/SmartArt support the driving goal (then Phases 2-4 first, skipping most of Phase 1), or reducing duplication first?
