# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.

Each dated section below names the per-package releases cut in that run; the per-package
history lives in `packages/<name>/CHANGELOG.md`. Nothing has been published yet, so there are no
sections: the first release run writes them.
## 2026-10-03

_Releases: docx-viewer-mcp@0.1.2, docx-core@0.1.4, docx-react-viewer@0.4.1, docx-vue-viewer@0.4.1, docx-angular-viewer@0.4.1, docx-svelte-viewer@0.4.1, docx-solid-viewer@0.4.1, docx-vanilla-viewer@0.4.1_

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([e6c664c](https://github.com/ChristopherVR/docx-viewer/commit/e6c664ccc06d5e207e4fd7c5a455433ab8868976))

### Chores

- **repo:** Remove Codex co-author trailers (by @ChristopherVR) ([cfa6b9f](https://github.com/ChristopherVR/docx-viewer/commit/cfa6b9f7268bdc39593bdb87a9818a7c3d1c3da5))

## 2026-10-03

_Releases: docx-viewer-mcp@0.1.1_

### Features

- **release:** Add core backed Word MCP tools (by @ChristopherVR) ([ebae8c7](https://github.com/ChristopherVR/docx-viewer/commit/ebae8c7a59641b3b3a59dc41c193d77e303d0bb4))

### Bug Fixes

- **release:** Stage planned package versions and changelogs (by @ChristopherVR) ([652c149](https://github.com/ChristopherVR/docx-viewer/commit/652c1498f4fd939c81cee0c10b7b9ad5e0525944))

### Documentation

- Remove stale npm publication notices (by @ChristopherVR) ([ed78f9f](https://github.com/ChristopherVR/docx-viewer/commit/ed78f9fce9198cd92ce51c79b9d36e7f6a33f37c))

## 2026-10-03

_Releases: docx-core@0.1.3, docx-react-viewer@0.4.0, docx-vue-viewer@0.4.0, docx-angular-viewer@0.4.0, docx-svelte-viewer@0.4.0, docx-solid-viewer@0.4.0, docx-vanilla-viewer@0.4.0_

### Features

- **web-component:** Edit selected cell vertical alignment (by @ChristopherVR) ([f744de1](https://github.com/ChristopherVR/docx-viewer/commit/f744de14876ad306fd8791f8ff2e38b987741b0b))

### Bug Fixes

- **web-component:** Preserve formatting in clipboard commands (by @ChristopherVR) ([453a451](https://github.com/ChristopherVR/docx-viewer/commit/453a451c1ce24044a2896feb58704f7eb318fb5f))
- **web-component:** Map print clicks to visible glyph boundaries (by @ChristopherVR) ([fc4e047](https://github.com/ChristopherVR/docx-viewer/commit/fc4e04778bdb4f38d6b8ec54e693c5f4f0b08220))

### Documentation

- Record alignment clipboard and print cursor coverage (by @ChristopherVR) ([7ba78de](https://github.com/ChristopherVR/docx-viewer/commit/7ba78de4f5d450354b55f302302a000e0af7f151))

### Dependencies

- **deps:** Consume shared alignment and layout source ranges (by @ChristopherVR) ([81b8c51](https://github.com/ChristopherVR/docx-viewer/commit/81b8c5148f139f9e8eadde14491eb167b042072b))

## 2026-10-03

_Releases: docx-core@0.1.2, docx-react-viewer@0.3.0, docx-vue-viewer@0.3.0, docx-angular-viewer@0.3.0, docx-svelte-viewer@0.3.0, docx-solid-viewer@0.3.0, docx-vanilla-viewer@0.3.0_

### Features

- **web-component:** Edit margins for selected table cells (by @ChristopherVR) ([52605c4](https://github.com/ChristopherVR/docx-viewer/commit/52605c40112d7c0f3a6c6c585c97dbabdafc7aed))
- **web-component:** Display and preserve imported equations (by @ChristopherVR) ([3286dc2](https://github.com/ChristopherVR/docx-viewer/commit/3286dc25f8258e17988a7a718b79c2c9466507f3))

### Bug Fixes

- **web-component:** Wrap scaled words across formatting runs (by @ChristopherVR) ([7821875](https://github.com/ChristopherVR/docx-viewer/commit/7821875536909eda84493337f23c8e0d861cc2ba))

### Documentation

- Record equation and cell margin support boundaries (by @ChristopherVR) ([a96105e](https://github.com/ChristopherVR/docx-viewer/commit/a96105ebb9edd360d24e0c6f9be92846a630e455))

### Dependencies

- **deps:** Consume shared equation and cell margin support (by @ChristopherVR) ([9e943fa](https://github.com/ChristopherVR/docx-viewer/commit/9e943fa61bb2f380c2a2cd1153bb1f0f6e276819))

## 2026-10-03

_Releases: docx-core@0.1.1, docx-react-viewer@0.2.0, docx-vue-viewer@0.2.0, docx-angular-viewer@0.2.0, docx-svelte-viewer@0.2.0, docx-solid-viewer@0.2.0, docx-vanilla-viewer@0.2.0_

### Features

- **web-component:** Improve file home and live editor options (by @ChristopherVR) ([7db1f08](https://github.com/ChristopherVR/docx-viewer/commit/7db1f08933994c46e98dab9b2464d2249391ac7b))
- **web-component:** Edit table shading with undo and preservation (by @ChristopherVR) ([1525735](https://github.com/ChristopherVR/docx-viewer/commit/15257358ed6653bb87047d501eec6558fb9b6b8e))

### Bug Fixes

- **build:** Ignore retired workspace artifacts in shared checks (by @ChristopherVR) ([6a7d531](https://github.com/ChristopherVR/docx-viewer/commit/6a7d531635ae91f74d155bf442347c5384e61b1c))

### Documentation

- Make AGENTS.md canonical and import it from CLAUDE.md (by @ChristopherVR) ([29e25e9](https://github.com/ChristopherVR/docx-viewer/commit/29e25e9a32791f30ad151299389365d70af0780b))
- Refresh word parity milestones and current evidence (by @ChristopherVR) ([ec8dca6](https://github.com/ChristopherVR/docx-viewer/commit/ec8dca6ae22b524c78646b9f6d0e53b3159bd352))

### Build & CI

- **release:** Run the release workflow hourly (by @ChristopherVR) ([960330a](https://github.com/ChristopherVR/docx-viewer/commit/960330a22ffce807c4005c6b09813afaa9f7ccd1))

### Dependencies

- **deps:** Consume the published office core table fixes (by @ChristopherVR) ([a7c1616](https://github.com/ChristopherVR/docx-viewer/commit/a7c16165fb91b88b4688a69fca13c4eecef3659e))
- **deps:** Align shared controls with the current office core (by @ChristopherVR) ([5f379cd](https://github.com/ChristopherVR/docx-viewer/commit/5f379cd71b4ad47fa8bb217ce6067af2c82da89a))

