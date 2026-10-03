# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs re-releases a package whenever any of its files, or an internal
dependency it ships against, changes, not only on conventional commits.

## [0.4.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-svelte-viewer@0.4.0) - 2026-10-03

### Features

- **web-component:** Edit selected cell vertical alignment (by @ChristopherVR) ([f744de1](https://github.com/ChristopherVR/docx-viewer/commit/f744de14876ad306fd8791f8ff2e38b987741b0b))

### Bug Fixes

- **web-component:** Preserve formatting in clipboard commands (by @ChristopherVR) ([453a451](https://github.com/ChristopherVR/docx-viewer/commit/453a451c1ce24044a2896feb58704f7eb318fb5f))
- **web-component:** Map print clicks to visible glyph boundaries (by @ChristopherVR) ([fc4e047](https://github.com/ChristopherVR/docx-viewer/commit/fc4e04778bdb4f38d6b8ec54e693c5f4f0b08220))

### Dependencies

- **deps:** Consume shared alignment and layout source ranges (by @ChristopherVR) ([81b8c51](https://github.com/ChristopherVR/docx-viewer/commit/81b8c5148f139f9e8eadde14491eb167b042072b))

## [0.3.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-svelte-viewer@0.3.0) - 2026-10-03

### Features

- **web-component:** Edit margins for selected table cells (by @ChristopherVR) ([52605c4](https://github.com/ChristopherVR/docx-viewer/commit/52605c40112d7c0f3a6c6c585c97dbabdafc7aed))
- **web-component:** Display and preserve imported equations (by @ChristopherVR) ([3286dc2](https://github.com/ChristopherVR/docx-viewer/commit/3286dc25f8258e17988a7a718b79c2c9466507f3))

### Bug Fixes

- **web-component:** Wrap scaled words across formatting runs (by @ChristopherVR) ([7821875](https://github.com/ChristopherVR/docx-viewer/commit/7821875536909eda84493337f23c8e0d861cc2ba))

### Dependencies

- **deps:** Consume shared equation and cell margin support (by @ChristopherVR) ([9e943fa](https://github.com/ChristopherVR/docx-viewer/commit/9e943fa61bb2f380c2a2cd1153bb1f0f6e276819))

## [0.2.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-svelte-viewer@0.2.0) - 2026-10-03

### Features

- **web-component:** Improve file home and live editor options (by @ChristopherVR) ([7db1f08](https://github.com/ChristopherVR/docx-viewer/commit/7db1f08933994c46e98dab9b2464d2249391ac7b))
- **web-component:** Edit table shading with undo and preservation (by @ChristopherVR) ([1525735](https://github.com/ChristopherVR/docx-viewer/commit/15257358ed6653bb87047d501eec6558fb9b6b8e))

### Dependencies

- **deps:** Consume the published office core table fixes (by @ChristopherVR) ([a7c1616](https://github.com/ChristopherVR/docx-viewer/commit/a7c16165fb91b88b4688a69fca13c4eecef3659e))
- **deps:** Align shared controls with the current office core (by @ChristopherVR) ([5f379cd](https://github.com/ChristopherVR/docx-viewer/commit/5f379cd71b4ad47fa8bb217ce6067af2c82da89a))

## [0.1.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-svelte-viewer@0.1.0) - 2026-10-02

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
- **editor:** Add Word-style window chrome and full-screen demo (by @ChristopherVR) ([7a5744c](https://github.com/ChristopherVR/docx-viewer/commit/7a5744c156bec56803dcd85adb53825641c9bbaf))
- **editor:** Add picture, hyperlink, character style and hidden-text commands (by @ChristopherVR) ([535758a](https://github.com/ChristopherVR/docx-viewer/commit/535758a75d88e2c886117f17d73a55e01fd6fed6))
- **editor:** Resize pictures, edit alt text, and honor toggle cancellation (by @ChristopherVR) ([5a3aa2d](https://github.com/ChristopherVR/docx-viewer/commit/5a3aa2d8eaad0569ced99a89aaad656913556675))
- **layout:** Show headers, footers and page numbers in Print Layout (by @ChristopherVR) ([825fbf5](https://github.com/ChristopherVR/docx-viewer/commit/825fbf5de71598bbdc4a2c60888805d9f8256ba7))
- Edit headers and footers in place (by @ChristopherVR) ([9699a8b](https://github.com/ChristopherVR/docx-viewer/commit/9699a8b50c4d962dc30af079d087c323f55c0862))
- Edit footnotes and endnotes in place (by @ChristopherVR) ([4976310](https://github.com/ChristopherVR/docx-viewer/commit/497631029e768d1f58761f26a6fdfc8b0e139884))
- Insert footnotes and endnotes (by @ChristopherVR) ([7ae74f4](https://github.com/ChristopherVR/docx-viewer/commit/7ae74f45712b6ab8ff962cd9bb088a38acac0ad5))
- Edit sections, page setup and section breaks (by @ChristopherVR) ([235d2f1](https://github.com/ChristopherVR/docx-viewer/commit/235d2f1740e19b283bcd807173c84cac881f0c60))
- **tables:** Render table style text formatting and cell margins (by @ChristopherVR) ([3cdfe41](https://github.com/ChristopherVR/docx-viewer/commit/3cdfe414077d94554f104882fb0887b58d2a6c15))
- **sections:** Edit different first page and page numbering (by @ChristopherVR) ([a6d12da](https://github.com/ChristopherVR/docx-viewer/commit/a6d12da51c6880eb8d5ee6dfb8df9da3493bf390))
- Pictures and links in headers, footers and notes (by @ChristopherVR) ([e4522f5](https://github.com/ChristopherVR/docx-viewer/commit/e4522f5e1de53afa78a71e418bec90edfc93aa94))
- **pictures:** Approximate floating picture wrapping (by @ChristopherVR) ([f904915](https://github.com/ChristopherVR/docx-viewer/commit/f9049151102532f15621e1435d429627fe9d6d64))
- **editor:** Ribbon targets the header, footer or note being edited (by @ChristopherVR) ([cd9d309](https://github.com/ChristopherVR/docx-viewer/commit/cd9d30953f667e414e62f833a72f2d1927b51bad))
- **sections:** Vertical alignment, odd/even breaks and odd/even headers (by @ChristopherVR) ([309b113](https://github.com/ChristopherVR/docx-viewer/commit/309b11359f14e3aa497b0a1ed8b600db35357a0f))
- **fields:** Update DATE and TIME fields in Print Layout (by @ChristopherVR) ([6890466](https://github.com/ChristopherVR/docx-viewer/commit/6890466525f91ae404b3ae998d69b82ff8228eb6))
- **pictures:** SVG pictures with PNG fallback (by @ChristopherVR) ([e583d8a](https://github.com/ChristopherVR/docx-viewer/commit/e583d8a4835571fb5d6d2ed8fb2a6525f61ecfda))
- **fields:** Editable complex and simple fields (by @ChristopherVR) ([43a7f7d](https://github.com/ChristopherVR/docx-viewer/commit/43a7f7d4da59210c7ecd80551402eb5f8a9b506a))
- **toc:** Insert and update tables of contents; paragraph tab stops (by @ChristopherVR) ([ce93a8d](https://github.com/ChristopherVR/docx-viewer/commit/ce93a8d529794878394d9b85aef10a16d5a75769))
- **print-layout:** Paginate inline pictures and position floating pictures (by @ChristopherVR) ([d68e089](https://github.com/ChristopherVR/docx-viewer/commit/d68e08945823b0c65badb45c855c2c7eccec3c25))
- **print-layout:** Style-resolved run metrics and metric-compatible fonts (by @ChristopherVR) ([f3f37fb](https://github.com/ChristopherVR/docx-viewer/commit/f3f37fb314cd5885c446aa4fcbf0cc7626e89d18))
- **comments:** Comment ranges spanning paragraphs; decimal comment ids (by @ChristopherVR) ([166e482](https://github.com/ChristopherVR/docx-viewer/commit/166e4825710cee33048ad581492c8632c0c8f661))
- **revisions:** Link tracked moves and keep their range markers (by @ChristopherVR) ([2ab59b2](https://github.com/ChristopherVR/docx-viewer/commit/2ab59b2f42f6eb877c34d3625b5448b4ae67548a))
- **print-layout:** Indents, tab stops with leaders, and list labels (by @ChristopherVR) ([a4f55a6](https://github.com/ChristopherVR/docx-viewer/commit/a4f55a6c29354ff452debae4af05fc37c8ad4140))
- **print-layout:** Header and footer pictures (by @ChristopherVR) ([b993c0d](https://github.com/ChristopherVR/docx-viewer/commit/b993c0dd088a4d114f8649eba3e8ffc5cd1374b1))
- **toc:** Hyperlinked entries with _Toc bookmarks and PAGEREF fields (by @ChristopherVR) ([5254ad9](https://github.com/ChristopherVR/docx-viewer/commit/5254ad96f008b9f7e78972052c680e0f45c94e3d))
- **editor:** Tab stops and leaders on the editing surface (by @ChristopherVR) ([f8181a8](https://github.com/ChristopherVR/docx-viewer/commit/f8181a887aba2a8e78b67706f06966cb2d67bee2))
- **track-changes:** Record moves; keep pasted formatting valid (by @ChristopherVR) ([5196ee7](https://github.com/ChristopherVR/docx-viewer/commit/5196ee7b9107d741251b11a1893115cba5939139))
- Word defaults for new documents; Print Layout tables (by @ChristopherVR) ([614d0fd](https://github.com/ChristopherVR/docx-viewer/commit/614d0fde2b600e601317c9aa1c226bda1e86b5f1))
- **print-layout:** Footnotes, endnotes and superscripts (by @ChristopherVR) ([2c48fc5](https://github.com/ChristopherVR/docx-viewer/commit/2c48fc5c2e8411f63c18cd06287426ab02bab363))
- **print-layout:** Baseline-aligned text (by @ChristopherVR) ([4ad14db](https://github.com/ChristopherVR/docx-viewer/commit/4ad14db7534cbc6935460f35b90e3140a19a54f7))
- **tables:** Row heights, header rows, keep-together and default cell margins (by @ChristopherVR) ([788b74a](https://github.com/ChristopherVR/docx-viewer/commit/788b74a6358f2081844f056fa91813569078c06f))
- **pagination:** Keep-with-next, keep-lines, widow control, contextual spacing (by @ChristopherVR) ([55ece2b](https://github.com/ChristopherVR/docx-viewer/commit/55ece2b0b6823983ea00fc44767e249c680e648f))
- Paragraph borders and shading (by @ChristopherVR) ([d9af53f](https://github.com/ChristopherVR/docx-viewer/commit/d9af53f85c6e9592f5df4cf285a6e1f65e845942))
- Added schemas (by @ChristopherVR) ([13ee39c](https://github.com/ChristopherVR/docx-viewer/commit/13ee39cc509504dbcc53f2b37b1e710432e4f526))
- **web-component:** Design tokens, dark mode and focus ring (by @claude) ([65d4d97](https://github.com/ChristopherVR/docx-viewer/commit/65d4d97d9dac947cf9a1d84c80f53f5d7e993e47))
- **core:** Schema-typed parsing with generated ECMA-376 simple types (by @claude) ([5aeb7b5](https://github.com/ChristopherVR/docx-viewer/commit/5aeb7b552eab29c6f3cdd8f88756d3a6a752065e))
- **web-component:** Typed events, reflected attributes and shared binding keys (by @claude) ([8b3d7a2](https://github.com/ChristopherVR/docx-viewer/commit/8b3d7a2c72e23a458ff12ef42a3cb9ba8da7b05d))
- **core:** Validate the model before save; write table cell props, exact jc and unique docPr ids (by @claude) ([fa24aac](https://github.com/ChristopherVR/docx-viewer/commit/fa24aacd9341604fa357c78a0a0d7ef34e4c134c))
- **web-component:** Page thumbnail rail, save/dirty API and ribbon customisation (by @claude) ([0239e27](https://github.com/ChristopherVR/docx-viewer/commit/0239e275630bdf4129de81febf143fcaf037846a))
- **web-component:** Central keyboard shortcuts, shortcut help and context menu (by @claude) ([581369f](https://github.com/ChristopherVR/docx-viewer/commit/581369feb861e700f81eec62f9a43d44dfc7afa0))
- **web-component:** Stable kebab-case ribbon action ids for hiddenActions (by @claude) ([a6441cb](https://github.com/ChristopherVR/docx-viewer/commit/a6441cb8d97e250bb3dc7c38c20879881841ed11))
- **web-component:** Per-locale strings with de, es and zh-CN, no hard-coded UI English (by @claude) ([7c609c1](https://github.com/ChristopherVR/docx-viewer/commit/7c609c1b1e9dab547a0f19af761915ce33a0c7a1))
- **core:** Create header/footer parts on save and write paragraph shading and borders (by @ChristopherVR) ([ec49241](https://github.com/ChristopherVR/docx-viewer/commit/ec492419dc52e83bbfda5866dac8f964111a153b))
- **web-component:** Word-style ribbon across every tab, dialogs and real commands (by @ChristopherVR) ([94a776b](https://github.com/ChristopherVR/docx-viewer/commit/94a776bc2d2aa4d4c40b7fcb2688b5fdefd2dc27))
- **web-component:** Word-style File backstage with Info, Save As, Print, Export and Options (by @ChristopherVR) ([1f7c3c0](https://github.com/ChristopherVR/docx-viewer/commit/1f7c3c027c9efff0c4d6cb55164d936cd11332f1))
- **web-component:** Word-style Styles gallery and Editing group on Home (by @ChristopherVR) ([df9232e](https://github.com/ChristopherVR/docx-viewer/commit/df9232e012aea84701f63d12c8d9440c1dd6fc9c))
- **web-component:** Home > Sort for selected paragraphs (by @ChristopherVR) ([9498d92](https://github.com/ChristopherVR/docx-viewer/commit/9498d929bcaae303adf6e6b5eea97df2c584bff6))
- **web-component:** Insert tab gets a table size picker, Bookmark dialog, Cover Page and Horizontal Line (by @ChristopherVR) ([a5009d7](https://github.com/ChristopherVR/docx-viewer/commit/a5009d753d0a75c5978ec00dc96d5e25fa4e1387))
- **web-component:** Page Setup dialog on Layout and outstanding-work update (by @ChristopherVR) ([cf7b815](https://github.com/ChristopherVR/docx-viewer/commit/cf7b8153510d728b0c36f037f84d2d52009ab678))
- **web-component:** View > Ruler with margins, inch ticks and indent markers (by @ChristopherVR) ([ee075dc](https://github.com/ChristopherVR/docx-viewer/commit/ee075dc997a90e490e73ce97ff58d9414ba853ed))
- **web-component:** References tab gets TOC depth and removal, captions and cross-references (by @ChristopherVR) ([49dcb35](https://github.com/ChristopherVR/docx-viewer/commit/49dcb357f772beb0671118ed2dd892ee7cae1f52))
- **web-component:** Review comments group gains Delete, Previous and Next (by @ChristopherVR) ([9921a8e](https://github.com/ChristopherVR/docx-viewer/commit/9921a8e732cdf017ca9fd89774624990ad84bdd4))
- **core:** Model page background and hyphenation settings (by @ChristopherVR) ([65e12fc](https://github.com/ChristopherVR/docx-viewer/commit/65e12fcba549494bff6c8d21782e96b9ccf26fb2))
- **web-component:** Ribbon parity pass with galleries, colour pickers, KeyTips, overflow and Zoom dialog (by @ChristopherVR) ([6555c45](https://github.com/ChristopherVR/docx-viewer/commit/6555c450267346b295995eea353b611e7c57217f))
- **web-component:** Margins gallery gains Custom Margins (by @ChristopherVR) ([904fd50](https://github.com/ChristopherVR/docx-viewer/commit/904fd50793d90c3397c4f84963a8b86f0620193e))
- **web-component:** Drag Ruler indent markers (by @ChristopherVR) ([1d6faac](https://github.com/ChristopherVR/docx-viewer/commit/1d6faacd6225604a9b4ec224c04eeae2f650a292))
- Insert > Drop Cap with framePr round trip and layout folding (by @ChristopherVR) ([cabd832](https://github.com/ChristopherVR/docx-viewer/commit/cabd8327d111f8d86a0a0294fb476b5e8510352d))
- Read text boxes and show their text read-only (by @ChristopherVR) ([f9e3c56](https://github.com/ChristopherVR/docx-viewer/commit/f9e3c565134048bb5b9082a3b8ef286534cb8d74))
- Home > Multilevel List with legal and outline styles (by @ChristopherVR) ([c113f64](https://github.com/ChristopherVR/docx-viewer/commit/c113f64cfdad3df8ae1e6bd1e7259680d3bd63f7))
- **web-component:** Borders and Shading dialog (by @ChristopherVR) ([ea1d591](https://github.com/ChristopherVR/docx-viewer/commit/ea1d591038fa646df4987d2f81000859e88819a3))
- Layout > Line Numbers with Print Layout rendering (by @ChristopherVR) ([881da88](https://github.com/ChristopherVR/docx-viewer/commit/881da88edbfb0d5592ef4d31819164c54a2cff22))
- **web-component:** References > Update Fields for SEQ, REF and PAGEREF (by @ChristopherVR) ([ff8c082](https://github.com/ChristopherVR/docx-viewer/commit/ff8c082167f909749d28a371721ae89847fb221a))
- Marker font controls for Define New Multilevel List (by @ChristopherVR) ([ef462cf](https://github.com/ChristopherVR/docx-viewer/commit/ef462cfe43fc24f62a1f65a0313967127a8ac23f))
- Heading-linked styles in Define New Multilevel List (by @ChristopherVR) ([c89e0e0](https://github.com/ChristopherVR/docx-viewer/commit/c89e0e001222c53b2134b8be7f8b3d2e07ca5c36))
- **web-component:** Heading-linked outlines in the Multilevel List gallery (by @ChristopherVR) ([40258a3](https://github.com/ChristopherVR/docx-viewer/commit/40258a3ace72df7c0e040fd7fd55b34222a8a2be))
- Editable document properties (title, subject, author, tags, comments) (by @ChristopherVR) ([3f53c23](https://github.com/ChristopherVR/docx-viewer/commit/3f53c233dd20c7c961a4137779ee3098e203aa84))
- Explicit tab stop per level in Define New Multilevel List (by @ChristopherVR) ([e7667ad](https://github.com/ChristopherVR/docx-viewer/commit/e7667ad28995dafc80e4a18f855f4058b663d81b))
- References > Table of Figures with update, and fix escaped field tests (by @ChristopherVR) ([4b239e2](https://github.com/ChristopherVR/docx-viewer/commit/4b239e2e2bf0759ba84cb27956969f7a80d53d9f))
- **web-component:** Update Fields also rebuilds the TOC and tables of figures (by @ChristopherVR) ([c56b3ae](https://github.com/ChristopherVR/docx-viewer/commit/c56b3ae06e1d42f2440b715aa1d5b182eda8516e))
- Table and cell border editor (Home > Borders and Borders and Shading Apply to) (by @ChristopherVR) ([2597e65](https://github.com/ChristopherVR/docx-viewer/commit/2597e65dc1bb083222922d89c68492c345264955))
- Layout > Page Borders (model, pgBorders write, Print Layout drawing, dialog) (by @ChristopherVR) ([99c4159](https://github.com/ChristopherVR/docx-viewer/commit/99c41590fac6aa108a518f7682d56428b8dd6824))
- Insert > Text Box (inline, editable), plus text-box line display and image run merge fixes (by @ChristopherVR) ([4d76849](https://github.com/ChristopherVR/docx-viewer/commit/4d768492d4ec50661356d63ff1431d0397cf77dc))
- Layout > Watermark (text watermark VML in headers, Print Layout layer, dialog) (by @ChristopherVR) ([21ec767](https://github.com/ChristopherVR/docx-viewer/commit/21ec7671afec67ba76f8da9b56d5059594591269))
- **web-component:** Inserting a caption also refreshes REF results that quote captions (by @ChristopherVR) ([c86273a](https://github.com/ChristopherVR/docx-viewer/commit/c86273a6959a63e2a08ad9bde21146fea858b500))
- **bindings:** Re-export the document model from every editor package (by @ChristopherVR) ([abf4e79](https://github.com/ChristopherVR/docx-viewer/commit/abf4e796049ab6584047c0cb66ba6c0af5396259))
- **web-component:** Display SmartArt from its cached drawing (by @ChristopherVR) ([4d5424c](https://github.com/ChristopherVR/docx-viewer/commit/4d5424cda33dccf12f009d33d2115fbbc74729b6))

### Bug Fixes

- **tables:** Keep cell formatting on edit, render table styles, write valid new tables (by @ChristopherVR) ([bf63864](https://github.com/ChristopherVR/docx-viewer/commit/bf63864b2f04487832e2ab366a9fd9385cba81dd))
- **notes:** Show note references in the document's number format (by @ChristopherVR) ([857fde5](https://github.com/ChristopherVR/docx-viewer/commit/857fde5c5908c01fa886420b33b3dc2faa6212fa))
- **revisions:** Save editor revision ids as decimals (by @ChristopherVR) ([bebb185](https://github.com/ChristopherVR/docx-viewer/commit/bebb1859f05b11a7528e9065cd061645a7dc058e))
- **formatting:** Word toggle-property semantics and explicit offs (by @ChristopherVR) ([84ad25d](https://github.com/ChristopherVR/docx-viewer/commit/84ad25d34f2e86460451bbaf05c0e9c2d9b1dcae))
- **core:** Write schema-ordered numbering, sectPr and drawing extents; make schema test real (by @claude) ([2b95fb7](https://github.com/ChristopherVR/docx-viewer/commit/2b95fb7ec32a2eae925d29f06d1302dfd059e785))
- **web-component:** Fit ribbon labels per locale, split tab keys, relocalize presence (by @claude) ([0881cfc](https://github.com/ChristopherVR/docx-viewer/commit/0881cfcfaa6d1c7dcd995e74a33fa37f4187dfc0))
- **packaging:** Viewer Svelte component imports bindings helpers; typed svelte.d.ts (by @claude) ([ac57529](https://github.com/ChristopherVR/docx-viewer/commit/ac575296c0d13efb5ac4625c34ffd136b1e6c603))
- **web-component:** Keep Define New Multilevel List within 720px with marker font row (by @ChristopherVR) ([cfee3e7](https://github.com/ChristopherVR/docx-viewer/commit/cfee3e78ef6812f44c204a4f5b844cc7f0f066e4))

### Other

- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([b4d3e44](https://github.com/ChristopherVR/docx-viewer/commit/b4d3e44463195e949316d3ac6df7bf9137485b4a))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([3cddda1](https://github.com/ChristopherVR/docx-viewer/commit/3cddda15ca30ac85612049ae80d63d947a230470))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([98b3036](https://github.com/ChristopherVR/docx-viewer/commit/98b3036664d3dedccfeb36dcbd9c99e7826795ec))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([52332bf](https://github.com/ChristopherVR/docx-viewer/commit/52332bf134d1c30a6ba63f7c61abbd79dad4b053))

### Refactor

- **core:** Remove writer helpers superseded by write-inline (by @ChristopherVR) ([f400024](https://github.com/ChristopherVR/docx-viewer/commit/f40002444f3c1dc14f1272e967c03c06865dabaa))
- **editor:** Move parts, page setup and list commands out of the component (by @ChristopherVR) ([bc8da60](https://github.com/ChristopherVR/docx-viewer/commit/bc8da6060474a78300616d83a55ec58c1d07b04a))
- Split ribbon, schema, model, adapter and collaboration modules (by @ChristopherVR) ([a3a1aee](https://github.com/ChristopherVR/docx-viewer/commit/a3a1aee3bf584c448f5de260b24a584dce346a60))
- Split range markers and body editor plugins into modules (by @ChristopherVR) ([7490a17](https://github.com/ChristopherVR/docx-viewer/commit/7490a17239c905b80c290b69402836fb704f50e6))
- **core:** Split oversized model and write modules (by @claude) ([b821cb6](https://github.com/ChristopherVR/docx-viewer/commit/b821cb64068d2b1887576770ea63de1fdcfbbc74))
- **web-component:** Split run mark construction out of run-adapter (by @claude) ([d4a2adf](https://github.com/ChristopherVR/docx-viewer/commit/d4a2adfe007d805912fa96bdf35e59d9dcc59243))
- **core,legacy:** Make source safe under noUncheckedIndexedAccess (by @claude) ([78b75e7](https://github.com/ChristopherVR/docx-viewer/commit/78b75e79a22e6d5e3aa0f3d9ebe068350143d74b))
- **web-component:** Make UI packages safe under noUncheckedIndexedAccess (by @claude) ([ec4a233](https://github.com/ChristopherVR/docx-viewer/commit/ec4a233783fe6066f1e8f814e671fb15638b8d82))
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
- Depend on the unscoped ooxml-core and ooxml-ui (by @ChristopherVR) ([c5be466](https://github.com/ChristopherVR/docx-viewer/commit/c5be466453eff0b35474b076558ba437f333b13e))

### Styling

- Apply oxfmt across the workspace (by @ChristopherVR) ([e9b3725](https://github.com/ChristopherVR/docx-viewer/commit/e9b3725582062b8140f76f7019912961b81b7ce8))
- **demo:** Widen style pickers and add a favicon (by @ChristopherVR) ([48c0044](https://github.com/ChristopherVR/docx-viewer/commit/48c0044ba1dadff52cd37460b440d1d9d6805ed1))

### Chores

- **types:** Enable exactOptionalPropertyTypes (by @claude) ([6f46221](https://github.com/ChristopherVR/docx-viewer/commit/6f4622160f37ae94a3bb5fa50710456b798dcaab))
- **types:** Enable noUncheckedIndexedAccess; keep test-support out of release builds (by @claude) ([203b289](https://github.com/ChristopherVR/docx-viewer/commit/203b289f0e8ca5d877cc6ff990ebc35755a853a3))
- Renamed ooxml packages [ci-skip] (by @ChristopherVR) ([3e75728](https://github.com/ChristopherVR/docx-viewer/commit/3e757282589b0327b90657d0d4c4e9350ce5025a))


