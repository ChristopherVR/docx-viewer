import type { LandingCopy } from './types';

export const en: LandingCopy = {
	hero: {
		kicker: 'Early implementation · Apache-2.0 · TypeScript',
		titleTop: '.docx editing,',
		titleAccent: 'made embeddable.',
		sub: 'A browser Word editor for React, Vue, Angular, Svelte, Solid, and vanilla JavaScript. One document model and one ProseMirror-backed <docx-editor> web component do the work; the framework adapters only wire up lifecycle and events.',
		start: { text: 'Get started', href: '/architecture' },
		demo: 'Live demo',
		scroll: 'Scroll',
		frameCaption: 'Sample document.docx · live in the browser',
		frameTry: 'Try it',
		frameAlt: 'The docx-viewer editor showing a sample document with the Word-style ribbon',
		copyLabel: 'Copy',
		copiedLabel: 'Copied',
		entryLabel: 'import from',
		notPublished:
			'Not on npm yet. The Word packages are unpublished, so these are the intended import paths. Build from the repository to try them today.',
	},
	features: {
		kicker: 'Features',
		title: 'What exists today.',
		items: [
			{
				title: 'One shared editor',
				copy: 'A Word-style ribbon, title bar, File backstage, and status bar live in one web component, with undo/redo, find and replace, zoom, and a read-only mode. Bindings add no UI of their own.',
				link: { text: 'Architecture', href: '/architecture' },
			},
			{
				title: 'Print Layout',
				copy: 'A paginated view with line breaking, widow/orphan and keep rules, columns, table row splitting, repeating headers, footnote placement, and printing. It is this editor’s own approximation, not Word’s pagination.',
				link: { text: 'Support roadmap', href: '/parity-roadmap' },
			},
			{
				title: 'Document structure',
				copy: 'Styles, numbered and bulleted lists, tables, sections and columns, headers and footers, footnotes, and fields including TOC. Inline text boxes, a header watermark, and page borders can be inserted.',
				link: { text: 'Outstanding work', href: '/outstanding-work' },
			},
			{
				title: 'Tracked changes and comments',
				copy: 'Track Changes mode, accept and reject, next and previous revision, and a comments pane with replies and resolve. Comments are written back to comments.xml.',
				link: { text: 'Editing text', href: '/editing' },
			},
			{
				title: 'Careful preservation',
				copy: 'A no-op save returns the original bytes. Supported edits preserve package parts outside the model, and edits that would damage unsupported content are rejected instead of dropped.',
				link: { text: 'Model and saving', href: '/architecture' },
			},
			{
				title: 'Legacy .doc',
				copy: 'Word 97-2003 files load through the shared ole2 codecs. Main-body text imports, and constrained edits to existing paragraphs save back as .doc. Formatting and structural changes are rejected.',
				link: { text: 'Support roadmap', href: '/parity-roadmap' },
			},
			{
				title: 'Collaboration protocol',
				copy: 'A transport-neutral protocol exchanges validated ProseMirror steps through an authority, with transient presence. Your application owns networking, identity, permissions, and storage.',
				link: { text: 'Collaboration', href: '/collaboration' },
			},
			{
				title: 'Localization',
				copy: 'The editor interface ships in English, French, German, Spanish, and Simplified Chinese through one locale option. It translates interface text only, never document content.',
				link: { text: 'Bindings', href: '/bindings' },
			},
		],
	},
	status: {
		kicker: 'Status',
		title: 'Honest about what is not done.',
		copy: 'This is an early implementation. It is not Microsoft Word parity: pagination is approximate, many features are not rendered with Word fidelity, password-protected files are unsupported, and Export DOCX cannot recover content the model omits. The support roadmap and outstanding-work pages list the gaps in detail.',
		link: { text: 'Support roadmap', href: '/parity-roadmap' },
		roadmapTitle: 'Roadmap note',
		roadmapCopy:
			'Shared OOXML logic is being moved into the private @christophervr/ooxml-core package, so this repository can end up holding only the UI. That migration is in progress and is not a shipped feature.',
		roadmapLink: { text: 'The ooxml-core plan', href: '/ooxml-core-plan' },
	},
	quickstart: {
		kicker: 'Getting started',
		title: 'Mount an editor in a few lines.',
		copy: 'Every adapter mounts the same <docx-editor>. Pass a document model, listen for changes, and give the container a height. The snippets show the intended API.',
		docsLabel: 'Framework guide',
		buildTitle: 'Build from source',
		buildCopy:
			'The Word packages are not published to npm yet, so clone the repository and run the demo or the packages from there. Bun is required.',
		buildCommands: 'bun install\nbun run demo',
	},
	demos: {
		kicker: 'Live demo',
		title: 'Try it right here.',
		copy: 'This is the real editor running in your browser: the demo app built from this repository for each framework adapter, embedded live. Open the sample document, or split the view to see two local peers share one document.',
		frameworkLabel: 'Framework',
		soloTab: 'Editor',
		collabTab: 'Collaboration',
		guestPicker: 'Guest',
		load: 'Load the live demo',
		loading: 'Loading the live editor',
		openFull: 'Open full app',
		hostLabel: 'Host',
		guestLabel: 'Guest',
		soloHint:
			'Use “Open the sample document” in the demo, or drop in a DOCX or DOC file of your own. Switching frameworks starts a fresh editor.',
		collabHint:
			'Two local peers share an in-memory authority. Pause delivery to try concurrent edits. This demo has no network backend; production transport is the host application’s job.',
	},
	faq: {
		kicker: 'FAQ',
		title: 'Common questions.',
		items: [
			{
				q: 'Can I install it from npm?',
				a: 'Not yet. The Word packages are not published, and the release workflow is deliberately not being run. Clone the repository and build it; the import paths shown on this site are the intended API.',
				link: { text: 'Release policy', href: '/releasing' },
			},
			{
				q: 'Is it free to use commercially?',
				a: 'The repository is Apache-2.0 licensed.',
			},
			{
				q: 'Is this Word parity?',
				a: 'No. Pagination is an approximation, many features are not rendered with Word fidelity, and Export DOCX creates a new document from supported visible content. Parity is tracked separately for import, layout, editing, export, preservation, and accessibility.',
				link: { text: 'Support roadmap', href: '/parity-roadmap' },
			},
			{
				q: 'Do edited files reopen in Word?',
				a: 'Supported edits are written back into the original package, and unsafe edits reject instead of silently degrading it. That is a design goal checked by regression tests, not a claim of lossless export.',
				link: { text: 'Architecture', href: '/architecture' },
			},
			{
				q: 'Which frameworks are supported?',
				a: 'React, Vue 3, Angular, Svelte 5, SolidJS, and plain JavaScript. They share one editor and model; each adapter only handles mounting, property updates, and event forwarding.',
				link: { text: 'Bindings', href: '/bindings' },
			},
			{
				q: 'Does collaboration need a server?',
				a: 'The protocol is transport-neutral and ships an in-memory reference authority. Networking, authentication, persistence, and permissions are your application’s responsibility.',
				link: { text: 'Collaboration', href: '/collaboration' },
			},
			{
				q: 'Can it open old .doc files?',
				a: 'Yes, with limits. Main-body text imports and constrained edits to existing paragraphs save back as .doc; formatting and structural changes are rejected.',
			},
			{
				q: 'What about password-protected files?',
				a: 'Unsupported today.',
			},
		],
	},
	finale: {
		kicker: 'Get started',
		title: '.docx in. .docx out.',
		sub: 'Read the architecture, pick your framework guide, and try the demo with a document of your own. Apache-2.0 licensed, strict TypeScript, and honest about the gaps.',
		quick: { text: 'Read the guide', href: '/architecture' },
		github: 'View on GitHub',
		columns: [
			{
				title: 'Product',
				links: [
					{
						text: 'Live demo',
						href: 'https://christophervr.github.io/docx-viewer/demo/',
						external: true,
					},
					{ text: 'Framework bindings', href: '/bindings' },
					{ text: 'Collaboration', href: '/collaboration' },
					{ text: 'Releases', href: '/releasing' },
				],
			},
			{
				title: 'Docs',
				links: [
					{ text: 'Architecture', href: '/architecture' },
					{ text: 'Editing text', href: '/editing' },
					{ text: 'Support roadmap', href: '/parity-roadmap' },
					{ text: 'Outstanding work', href: '/outstanding-work' },
				],
			},
			{
				title: 'Community',
				links: [
					{ text: 'GitHub', href: 'https://github.com/ChristopherVR/docx-viewer', external: true },
					{
						text: 'Issues',
						href: 'https://github.com/ChristopherVR/docx-viewer/issues',
						external: true,
					},
					{
						text: 'License',
						href: 'https://github.com/ChristopherVR/docx-viewer/blob/main/LICENSE',
						external: true,
					},
				],
			},
		],
		bottomLeft: '© 2026 ChristopherVR · Apache-2.0',
		bottomRight: 'docx-viewer · an early Word editor for the web',
	},
};
