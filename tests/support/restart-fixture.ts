// Copy of the fixture in ooxml-core (src/docx/test-support); the browser specs cannot import test-only code from the published package.
import JSZip from 'jszip';

export const restartSequence = [0, 1, 2, 2, 0, 2, 1, 2];
export const restartCases = [
	{ name: 'default', restart: undefined, expected: [1, 1, 1, 2, 2, 1, 2, 1] },
	{ name: 'never', restart: 0, expected: [1, 1, 1, 2, 2, 3, 2, 4] },
	{ name: 'after-level-one', restart: 1, expected: [1, 1, 1, 2, 2, 1, 2, 2] },
] as const;

/** Portable package used both in browser contracts and independent Word COM comparison. */
export async function restartFixture(restart: number | undefined, ancestorNever = false) {
	const zip = new JSZip();
	const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
	const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/numbering" Target="numbering.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${restartSequence.map((level, i) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item ${i}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/numbering.xml',
		`<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${[0, 1, 2].map((level) => `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>${level === 2 && restart !== undefined ? `<w:lvlRestart w:val="${restart}"/>` : level === 1 && ancestorNever ? '<w:lvlRestart w:val="0"/>' : ''}<w:lvlText w:val="%${level + 1}."/></w:lvl>`).join('')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
