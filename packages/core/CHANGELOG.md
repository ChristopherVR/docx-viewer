# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs re-releases a package whenever any of its files, or an internal
dependency it ships against, changes, not only on conventional commits.

## [0.1.3](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.3) - 2026-10-03

### Dependencies

- **deps:** Consume shared alignment and layout source ranges (by @ChristopherVR) ([81b8c51](https://github.com/ChristopherVR/docx-viewer/commit/81b8c5148f139f9e8eadde14491eb167b042072b))

## [0.1.2](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.2) - 2026-10-03

### Dependencies

- **deps:** Consume shared equation and cell margin support (by @ChristopherVR) ([9e943fa](https://github.com/ChristopherVR/docx-viewer/commit/9e943fa61bb2f380c2a2cd1153bb1f0f6e276819))

## [0.1.1](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.1) - 2026-10-03

### Features

- Add shared Word editor and multi-framework demos (by @ChristopherVR) ([77ea606](https://github.com/ChristopherVR/docx-viewer/commit/77ea606625752f976ed426967614b3bf31e13403))
- **editor:** Add Word run formatting and safe table editing (by @ChristopherVR) ([c06bbaf](https://github.com/ChristopherVR/docx-viewer/commit/c06bbaf25d6b5baa3e57dbcd2ee1bc88e0be1b16))
- **editor:** Add line spacing and hard line breaks (by @ChristopherVR) ([b414ec4](https://github.com/ChristopherVR/docx-viewer/commit/b414ec46857e1309ff31aa7603bf033874301c84))
- **editor:** Add collaboration, multilingual editing and search (by @ChristopherVR) ([da68728](https://github.com/ChristopherVR/docx-viewer/commit/da687283507df6e7aa6886fe866405c454aaa16c))
- Add list numbering, sections, breaks, headers/footers and notes (by @ChristopherVR) ([6a5462f](https://github.com/ChristopherVR/docx-viewer/commit/6a5462f973d64dfa3268baba4fb0c73f3fa76d5a))
- **layout:** Add pagination engine and Print Layout view (by @ChristopherVR) ([c30c6e1](https://github.com/ChristopherVR/docx-viewer/commit/c30c6e1d456e2357239468e1946d4577e544fb9c))
- **review:** Add tracked changes and comments (by @ChristopherVR) ([4d7a1e5](https://github.com/ChristopherVR/docx-viewer/commit/4d7a1e52e0417ccff8319355dd244da30af2b40e))
- **core:** Add theme, character styles and table fidelity (by @ChristopherVR) ([8bf918c](https://github.com/ChristopherVR/docx-viewer/commit/8bf918cbd8f2115f668d0446acbe146253249eeb))
- Add inline pictures, hyperlinks and bookmarks (by @ChristopherVR) ([7e0ff2b](https://github.com/ChristopherVR/docx-viewer/commit/7e0ff2bf6c6428057d4ef8eaa0bf2ccb6c4e68f5))
- **layout:** Show headers, footers and page numbers in Print Layout (by @ChristopherVR) ([825fbf5](https://github.com/ChristopherVR/docx-viewer/commit/825fbf5de71598bbdc4a2c60888805d9f8256ba7))
- Edit headers and footers in place (by @ChristopherVR) ([9699a8b](https://github.com/ChristopherVR/docx-viewer/commit/9699a8b50c4d962dc30af079d087c323f55c0862))
- Edit footnotes and endnotes in place (by @ChristopherVR) ([4976310](https://github.com/ChristopherVR/docx-viewer/commit/497631029e768d1f58761f26a6fdfc8b0e139884))
- Insert footnotes and endnotes (by @ChristopherVR) ([7ae74f4](https://github.com/ChristopherVR/docx-viewer/commit/7ae74f45712b6ab8ff962cd9bb088a38acac0ad5))
- Edit sections, page setup and section breaks (by @ChristopherVR) ([235d2f1](https://github.com/ChristopherVR/docx-viewer/commit/235d2f1740e19b283bcd807173c84cac881f0c60))
- **sections:** Edit different first page and page numbering (by @ChristopherVR) ([a6d12da](https://github.com/ChristopherVR/docx-viewer/commit/a6d12da51c6880eb8d5ee6dfb8df9da3493bf390))
- Pictures and links in headers, footers and notes (by @ChristopherVR) ([e4522f5](https://github.com/ChristopherVR/docx-viewer/commit/e4522f5e1de53afa78a71e418bec90edfc93aa94))
- **pictures:** Approximate floating picture wrapping (by @ChristopherVR) ([f904915](https://github.com/ChristopherVR/docx-viewer/commit/f9049151102532f15621e1435d429627fe9d6d64))
- **sections:** Vertical alignment, odd/even breaks and odd/even headers (by @ChristopherVR) ([309b113](https://github.com/ChristopherVR/docx-viewer/commit/309b11359f14e3aa497b0a1ed8b600db35357a0f))
- **fields:** Update DATE and TIME fields in Print Layout (by @ChristopherVR) ([6890466](https://github.com/ChristopherVR/docx-viewer/commit/6890466525f91ae404b3ae998d69b82ff8228eb6))
- **pictures:** SVG pictures with PNG fallback (by @ChristopherVR) ([e583d8a](https://github.com/ChristopherVR/docx-viewer/commit/e583d8a4835571fb5d6d2ed8fb2a6525f61ecfda))
- **fields:** Editable complex and simple fields (by @ChristopherVR) ([43a7f7d](https://github.com/ChristopherVR/docx-viewer/commit/43a7f7d4da59210c7ecd80551402eb5f8a9b506a))
- **toc:** Insert and update tables of contents; paragraph tab stops (by @ChristopherVR) ([ce93a8d](https://github.com/ChristopherVR/docx-viewer/commit/ce93a8d529794878394d9b85aef10a16d5a75769))
- **print-layout:** Paginate inline pictures and position floating pictures (by @ChristopherVR) ([d68e089](https://github.com/ChristopherVR/docx-viewer/commit/d68e08945823b0c65badb45c855c2c7eccec3c25))
- **comments:** Comment ranges spanning paragraphs; decimal comment ids (by @ChristopherVR) ([166e482](https://github.com/ChristopherVR/docx-viewer/commit/166e4825710cee33048ad581492c8632c0c8f661))
- **revisions:** Link tracked moves and keep their range markers (by @ChristopherVR) ([2ab59b2](https://github.com/ChristopherVR/docx-viewer/commit/2ab59b2f42f6eb877c34d3625b5448b4ae67548a))
- **print-layout:** Indents, tab stops with leaders, and list labels (by @ChristopherVR) ([a4f55a6](https://github.com/ChristopherVR/docx-viewer/commit/a4f55a6c29354ff452debae4af05fc37c8ad4140))
- **toc:** Hyperlinked entries with _Toc bookmarks and PAGEREF fields (by @ChristopherVR) ([5254ad9](https://github.com/ChristopherVR/docx-viewer/commit/5254ad96f008b9f7e78972052c680e0f45c94e3d))
- **track-changes:** Record moves; keep pasted formatting valid (by @ChristopherVR) ([5196ee7](https://github.com/ChristopherVR/docx-viewer/commit/5196ee7b9107d741251b11a1893115cba5939139))
- Word defaults for new documents; Print Layout tables (by @ChristopherVR) ([614d0fd](https://github.com/ChristopherVR/docx-viewer/commit/614d0fde2b600e601317c9aa1c226bda1e86b5f1))
- **print-layout:** Footnotes, endnotes and superscripts (by @ChristopherVR) ([2c48fc5](https://github.com/ChristopherVR/docx-viewer/commit/2c48fc5c2e8411f63c18cd06287426ab02bab363))
- **tables:** Row heights, header rows, keep-together and default cell margins (by @ChristopherVR) ([788b74a](https://github.com/ChristopherVR/docx-viewer/commit/788b74a6358f2081844f056fa91813569078c06f))
- **pagination:** Keep-with-next, keep-lines, widow control, contextual spacing (by @ChristopherVR) ([55ece2b](https://github.com/ChristopherVR/docx-viewer/commit/55ece2b0b6823983ea00fc44767e249c680e648f))
- Paragraph borders and shading (by @ChristopherVR) ([d9af53f](https://github.com/ChristopherVR/docx-viewer/commit/d9af53f85c6e9592f5df4cf285a6e1f65e845942))
- Added schemas (by @ChristopherVR) ([13ee39c](https://github.com/ChristopherVR/docx-viewer/commit/13ee39cc509504dbcc53f2b37b1e710432e4f526))
- **core:** Schema-typed parsing with generated ECMA-376 simple types (by @claude) ([5aeb7b5](https://github.com/ChristopherVR/docx-viewer/commit/5aeb7b552eab29c6f3cdd8f88756d3a6a752065e))
- **core:** Validate the model before save; write table cell props, exact jc and unique docPr ids (by @claude) ([fa24aac](https://github.com/ChristopherVR/docx-viewer/commit/fa24aacd9341604fa357c78a0a0d7ef34e4c134c))
- **core:** Create header/footer parts on save and write paragraph shading and borders (by @ChristopherVR) ([ec49241](https://github.com/ChristopherVR/docx-viewer/commit/ec492419dc52e83bbfda5866dac8f964111a153b))
- **core:** Model page background and hyphenation settings (by @ChristopherVR) ([65e12fc](https://github.com/ChristopherVR/docx-viewer/commit/65e12fcba549494bff6c8d21782e96b9ccf26fb2))
- Insert > Drop Cap with framePr round trip and layout folding (by @ChristopherVR) ([cabd832](https://github.com/ChristopherVR/docx-viewer/commit/cabd8327d111f8d86a0a0294fb476b5e8510352d))
- Read text boxes and show their text read-only (by @ChristopherVR) ([f9e3c56](https://github.com/ChristopherVR/docx-viewer/commit/f9e3c565134048bb5b9082a3b8ef286534cb8d74))
- Home > Multilevel List with legal and outline styles (by @ChristopherVR) ([c113f64](https://github.com/ChristopherVR/docx-viewer/commit/c113f64cfdad3df8ae1e6bd1e7259680d3bd63f7))
- Layout > Line Numbers with Print Layout rendering (by @ChristopherVR) ([881da88](https://github.com/ChristopherVR/docx-viewer/commit/881da88edbfb0d5592ef4d31819164c54a2cff22))
- Marker font controls for Define New Multilevel List (by @ChristopherVR) ([ef462cf](https://github.com/ChristopherVR/docx-viewer/commit/ef462cfe43fc24f62a1f65a0313967127a8ac23f))
- Heading-linked styles in Define New Multilevel List (by @ChristopherVR) ([c89e0e0](https://github.com/ChristopherVR/docx-viewer/commit/c89e0e001222c53b2134b8be7f8b3d2e07ca5c36))
- **web-component:** Heading-linked outlines in the Multilevel List gallery (by @ChristopherVR) ([40258a3](https://github.com/ChristopherVR/docx-viewer/commit/40258a3ace72df7c0e040fd7fd55b34222a8a2be))
- Editable document properties (title, subject, author, tags, comments) (by @ChristopherVR) ([3f53c23](https://github.com/ChristopherVR/docx-viewer/commit/3f53c233dd20c7c961a4137779ee3098e203aa84))
- Explicit tab stop per level in Define New Multilevel List (by @ChristopherVR) ([e7667ad](https://github.com/ChristopherVR/docx-viewer/commit/e7667ad28995dafc80e4a18f855f4058b663d81b))
- References > Table of Figures with update, and fix escaped field tests (by @ChristopherVR) ([4b239e2](https://github.com/ChristopherVR/docx-viewer/commit/4b239e2e2bf0759ba84cb27956969f7a80d53d9f))
- Table and cell border editor (Home > Borders and Borders and Shading Apply to) (by @ChristopherVR) ([2597e65](https://github.com/ChristopherVR/docx-viewer/commit/2597e65dc1bb083222922d89c68492c345264955))
- Layout > Page Borders (model, pgBorders write, Print Layout drawing, dialog) (by @ChristopherVR) ([99c4159](https://github.com/ChristopherVR/docx-viewer/commit/99c41590fac6aa108a518f7682d56428b8dd6824))
- Insert > Text Box (inline, editable), plus text-box line display and image run merge fixes (by @ChristopherVR) ([4d76849](https://github.com/ChristopherVR/docx-viewer/commit/4d768492d4ec50661356d63ff1431d0397cf77dc))
- Layout > Watermark (text watermark VML in headers, Print Layout layer, dialog) (by @ChristopherVR) ([21ec767](https://github.com/ChristopherVR/docx-viewer/commit/21ec7671afec67ba76f8da9b56d5059594591269))

### Bug Fixes

- **tables:** Keep cell formatting on edit, render table styles, write valid new tables (by @ChristopherVR) ([bf63864](https://github.com/ChristopherVR/docx-viewer/commit/bf63864b2f04487832e2ab366a9fd9385cba81dd))
- **revisions:** Save editor revision ids as decimals (by @ChristopherVR) ([bebb185](https://github.com/ChristopherVR/docx-viewer/commit/bebb1859f05b11a7528e9065cd061645a7dc058e))
- **formatting:** Word toggle-property semantics and explicit offs (by @ChristopherVR) ([84ad25d](https://github.com/ChristopherVR/docx-viewer/commit/84ad25d34f2e86460451bbaf05c0e9c2d9b1dcae))
- **core:** Write schema-ordered numbering, sectPr and drawing extents; make schema test real (by @claude) ([2b95fb7](https://github.com/ChristopherVR/docx-viewer/commit/2b95fb7ec32a2eae925d29f06d1302dfd059e785))
- **packaging:** Viewer Svelte component imports bindings helpers; typed svelte.d.ts (by @claude) ([ac57529](https://github.com/ChristopherVR/docx-viewer/commit/ac575296c0d13efb5ac4625c34ffd136b1e6c603))

### Other

- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([b4d3e44](https://github.com/ChristopherVR/docx-viewer/commit/b4d3e44463195e949316d3ac6df7bf9137485b4a))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([3cddda1](https://github.com/ChristopherVR/docx-viewer/commit/3cddda15ca30ac85612049ae80d63d947a230470))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([98b3036](https://github.com/ChristopherVR/docx-viewer/commit/98b3036664d3dedccfeb36dcbd9c99e7826795ec))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([52332bf](https://github.com/ChristopherVR/docx-viewer/commit/52332bf134d1c30a6ba63f7c61abbd79dad4b053))

### Refactor

- **core:** Remove writer helpers superseded by write-inline (by @ChristopherVR) ([f400024](https://github.com/ChristopherVR/docx-viewer/commit/f40002444f3c1dc14f1272e967c03c06865dabaa))
- Split ribbon, schema, model, adapter and collaboration modules (by @ChristopherVR) ([a3a1aee](https://github.com/ChristopherVR/docx-viewer/commit/a3a1aee3bf584c448f5de260b24a584dce346a60))
- Split range markers and body editor plugins into modules (by @ChristopherVR) ([7490a17](https://github.com/ChristopherVR/docx-viewer/commit/7490a17239c905b80c290b69402836fb704f50e6))
- **core:** Split oversized model and write modules (by @claude) ([b821cb6](https://github.com/ChristopherVR/docx-viewer/commit/b821cb64068d2b1887576770ea63de1fdcfbbc74))
- **core,legacy:** Make source safe under noUncheckedIndexedAccess (by @claude) ([78b75e7](https://github.com/ChristopherVR/docx-viewer/commit/78b75e79a22e6d5e3aa0f3d9ebe068350143d74b))
- **core:** Brand model measurements as Twips, SignedTwips and EighthPoints (by @claude) ([b4ee400](https://github.com/ChristopherVR/docx-viewer/commit/b4ee40048cb4d8a5f31ee26c5c23f7ff6cdbee89))
- **core:** Consume @christophervr/ooxml-units from the shared ooxml-core repo (by @ChristopherVR) ([3be6e4c](https://github.com/ChristopherVR/docx-viewer/commit/3be6e4c3a3e249f917fbf452c1ef2b2618f57216))
- **core:** Build the Word XML helpers on @christophervr/ooxml-xml (by @ChristopherVR) ([8795a34](https://github.com/ChristopherVR/docx-viewer/commit/8795a344f0656bcd810403927bf4a2c0518fac52))
- **core:** Consume relationship and content-type handling from ooxml-opc (by @ChristopherVR) ([5fbdea1](https://github.com/ChristopherVR/docx-viewer/commit/5fbdea1af1766ffba4b3580f50b735956591dc1f))
- Consume the single @christophervr/ooxml-core package (by @ChristopherVR) ([9fdcc2f](https://github.com/ChristopherVR/docx-viewer/commit/9fdcc2fc2a50dc59bf585e3b2138cc83e6047a26))
- **core:** Reduce docx-core to a thin re-export of @christophervr/ooxml-core/docx (by @ChristopherVR) ([76fb536](https://github.com/ChristopherVR/docx-viewer/commit/76fb536f2f9f6f9931226a10ac8b9aeef46c316f))
- **packages:** Publish one self-contained package per framework (by @ChristopherVR) ([f009c62](https://github.com/ChristopherVR/docx-viewer/commit/f009c6266315805753bcc55ce0688272396a8d9b))
- Adopt ooxml-core layout, load and collab areas (by @ChristopherVR) ([aa42730](https://github.com/ChristopherVR/docx-viewer/commit/aa427301f8af62707e50db6c915afa0b91b579a6))

### Documentation

- **schemas:** Add Ecma and W3C licence texts for the vendored XSDs (by @claude) ([210a1a5](https://github.com/ChristopherVR/docx-viewer/commit/210a1a576906172045c844c279c6e6e1ceda23c3))

### Testing

- Make browser runner resilient to missing pinned Chromium; document schema licence provenance (by @claude) ([6299813](https://github.com/ChristopherVR/docx-viewer/commit/6299813b28fa36f2de2078f1ab1c260194434187))
- **core:** Make tests safe under noUncheckedIndexedAccess (by @claude) ([ac9f888](https://github.com/ChristopherVR/docx-viewer/commit/ac9f8887cb56aad68b36d09fb05868d611cdf0c3))

### Build & CI

- Consume the published @christophervr/ooxml-core ^0.1.0 (by @ChristopherVR) ([38df96a](https://github.com/ChristopherVR/docx-viewer/commit/38df96a11de41072ece01648b8da118fe5729b7a))
- Publish under unscoped docx-* package names (by @ChristopherVR) ([44b369f](https://github.com/ChristopherVR/docx-viewer/commit/44b369ff2df6e34f210f7752ebd510df9464b405))

### Styling

- Apply oxfmt across the workspace (by @ChristopherVR) ([e9b3725](https://github.com/ChristopherVR/docx-viewer/commit/e9b3725582062b8140f76f7019912961b81b7ce8))

### Dependencies

- **deps:** Consume the published office core table fixes (by @ChristopherVR) ([a7c1616](https://github.com/ChristopherVR/docx-viewer/commit/a7c16165fb91b88b4688a69fca13c4eecef3659e))

### Chores

- **types:** Enable exactOptionalPropertyTypes (by @claude) ([6f46221](https://github.com/ChristopherVR/docx-viewer/commit/6f4622160f37ae94a3bb5fa50710456b798dcaab))
- **types:** Enable noUncheckedIndexedAccess; keep test-support out of release builds (by @claude) ([203b289](https://github.com/ChristopherVR/docx-viewer/commit/203b289f0e8ca5d877cc6ff990ebc35755a853a3))
- Renamed ooxml packages [ci-skip] (by @ChristopherVR) ([3e75728](https://github.com/ChristopherVR/docx-viewer/commit/3e757282589b0327b90657d0d4c4e9350ce5025a))


