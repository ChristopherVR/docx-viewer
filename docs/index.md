---
layout: home

hero:
  name: docx-viewer
  text: A shared foundation for Word editing on the web
  tagline: One document model and one editor, with thin adapters for the frameworks your app already uses.
  actions:
    - theme: brand
      text: Open the React demo
      link: /demo/
    - theme: alt
      text: Read the architecture
      link: /architecture

features:
  - title: Framework neutral
    details: A canonical document model, DOCX codec, and custom-element editor power every framework adapter.
  - title: Preserve carefully
    details: Unchanged documents retain their original bytes. Edited export covers the supported subset and reports known limits.
  - title: Try five demos
    details: Explore the shared editing surface mounted through React, Vue, Angular, Svelte, or vanilla JavaScript.
---

## Try the demos

Each route builds the same demo from `demos/demo-vanilla`, selecting one public framework adapter at build time.

The [coauthoring demo](/demo/collaboration.html) connects two editors through a local
in-memory authority. Pause delivery to try concurrent changes and resume to merge them.
Read the [collaboration integration guide](/collaboration) before connecting a server.

| Framework          | Demo                        | Guide                                      |
| ------------------ | --------------------------- | ------------------------------------------ |
| React              | [Open demo](/demo/)         | [React integration](/frameworks/react)     |
| Vue                | [Open demo](/demo-vue/)     | [Vue integration](/frameworks/vue)         |
| Angular            | [Open demo](/demo-angular/) | [Angular integration](/frameworks/angular) |
| Vanilla JavaScript | [Open demo](/demo-vanilla/) | [Vanilla integration](/frameworks/vanilla) |
| Svelte             | [Open demo](/demo-svelte/)  | [Svelte integration](/frameworks/svelte)   |

## Supported document subset

The editor currently handles paragraphs, direct text formatting, paragraph alignment, and simple table text. It provides a continuous editing surface rather than Word's pagination engine. Images, lists, headers, footers, fields, tracked changes, and other advanced features are not fully rendered or editable. Unsupported package content may remain in an unchanged document, but edits around it can affect preservation. DOC export supports a restricted paragraph/text path; DOCX export from the demo creates a new document from the visible model.

See the [support roadmap](/parity-roadmap) for the current limits and validation milestones.

## Architecture and ownership

The framework packages are thin lifecycle and event adapters. The web component owns editing and rendering, while the shared DOCX package owns canonical parsing and serialization. The separate document package detects DOCX and legacy DOC inputs. The [architecture guide](/architecture) describes these boundaries, and each [framework guide](/bindings) shows its public integration surface.

## Package entry points

The commands below list the intended public package names. The Word packages are not released yet; the current publishing scope is `@christophervr/ole2`. Use the source workspace for the Word editor until its packages are released.

```sh
npm install @christophervr/docx-core @christophervr/docx-document @christophervr/docx-web-component @christophervr/docx-bindings
```

Add the legacy adapter only when your application needs the supported `.doc` input path:

```sh
npm install @christophervr/docx-legacy
```

Use the framework-specific entry point documented in [framework bindings](/bindings). Package support remains an early editing subset; consult the roadmap before relying on advanced Word features.
