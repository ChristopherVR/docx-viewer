import { it } from 'vitest';
import {
	buildTableOfContents,
	createDocument,
	ensureListDefinition,
	saveDocx,
	tocBookmarks,
	tocEntries,
	withTocBookmarks,
	type DocumentModel,
} from './index.js';
import { packageSchemaErrors } from './test-support/schema-validation.js';

function kitchenSink(): DocumentModel {
	const model = createDocument();
	const rev = { author: 'Ada', date: '2024-01-01T00:00:00Z' };
	model.blocks = [
		{ type: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: 'Heading' }], keepNext: true },
		{
			type: 'paragraph',
			id: 'p1',
			align: 'justify',
			spacingBeforeTwips: 120,
			lineSpacingTwips: 360,
			lineSpacingRule: 'auto',
			indentLeftTwips: 720,
			hangingTwips: 360,
			tabStops: [{ posTwips: 4680, align: 'right', leader: 'dot' }],
			pageBreakBefore: true,
			direction: 'rtl',
			runs: [
				{ text: 'Bold', bold: true, italic: false, underline: true, strike: false, color: '#FF0000', fontSize: 14, fontFamily: 'Arial', highlight: 'yellow', verticalAlign: 'superscript', language: 'en-US', rtl: true, caps: true, smallCaps: false, vanish: false, characterSpacingTwips: 20, shadingFill: '#FFFF00', underlineStyle: 'double', underlineColor: '#00FF00', doubleStrike: false, style: 'Hyperlink' },
				{ text: 'link', link: { anchor: '_Toc100000000' } },
				{ text: '', break: 'page' },
				{ text: '\tafter tab\n' },
				{ text: 'inserted', revision: { ...rev, kind: 'insert', id: '5' } },
				{ text: 'deleted', revision: { ...rev, kind: 'delete', id: '6' } },
				{ text: 'moved', revision: { ...rev, kind: 'moveFrom', id: '7', move: { name: 'move1' } } },
				{ text: 'moved', revision: { ...rev, kind: 'moveTo', id: '8', move: { name: 'move1' } } },
				{ text: 'commented', commentIds: ['c1'] },
				{ text: '', noteReference: { kind: 'footnote', id: '1' } },
				{ text: '', fieldChar: 'begin' },
				{ text: '', fieldCode: ' PAGE ' },
				{ text: '', fieldChar: 'separate' },
				{ text: '1', field: { instr: 'PAGE' } },
				{ text: '', fieldChar: 'end' },
				{ text: 'Ann', field: { instr: 'AUTHOR', simple: true } },
			],
			markRevision: { ...rev, kind: 'insert', id: '9' },
		},
		{
			type: 'table',
			id: 't1',
			rows: [
				[
					{ paragraphs: [{ type: 'paragraph', id: 'c1p', runs: [{ text: 'A' }] }], gridSpan: 1, verticalAlign: 'center', shadingFill: 'D9D9D9', widthTwips: 2000 },
					{ paragraphs: [{ type: 'paragraph', id: 'c2p', runs: [{ text: 'B' }] }] },
				],
			],
		},
		{ type: 'paragraph', id: 'end', runs: [{ text: 'End' }] },
	];
	model.comments = [{ id: 'c1', author: 'Ada', text: 'Note this', resolved: true }];
	model.footnotes = [{ id: '1', blocks: [{ type: 'paragraph', id: 'n1', style: 'FootnoteText', runs: [{ text: '', noteMark: 'footnote', style: 'FootnoteReference' }, { text: ' Source' }] }] }];
	const withList = ensureListDefinition(model, 'bullet');
	return withList.model ?? model;
}

it('probe', async () => {
	const model = kitchenSink();
	const { bookmarks, added } = tocBookmarks(model, tocEntries(model));
	const marked = withTocBookmarks(model, added);
	marked.blocks.unshift(...buildTableOfContents(marked, { newId: (() => { let n = 0; return () => `toc${++n}`; })(), bookmarks, pageNumbers: new Map([['h1', '1']]) }));
	const errors = await packageSchemaErrors(await saveDocx(marked));
	for (const [part, list] of Object.entries(errors)) {
		console.log('PART', part, list.length);
		for (const line of [...new Set(list.map((e) => e.replace(/^part\.xml:\d+: Schemas validity error : /, '').replace(/\{http:\/\/schemas\.openxmlformats\.org\/wordprocessingml\/2006\/main\}/g, 'w:')))]) console.log('  ', line.slice(0, 260));
	}
}, 60000);
