// Copy of the fixture in ooxml-core (src/docx/test-support); the browser specs cannot import test-only code from the published package.
import JSZip from 'jszip';

export const headingSequence = [1, 2, 2, 1, 2];
export const headingLabels = ['1.', '1.1.', '1.2.', '2.', '2.1.'];

/** A real style-linked outline: styles carry numId only; levels carry pStyle. */
export async function headingNumberingFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
	const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
	const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="${office}/styles" Target="styles.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body>${headingSequence.map((level, index) => `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t>Heading ${index}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/styles.xml',
		`<w:styles xmlns:w="${w}">${[1, 2].map((level) => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="${level - 1}"/></w:pPr></w:style>`).join('')}</w:styles>`,
	);
	zip.file(
		'word/numbering.xml',
		`<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${[0, 1].map((level) => `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading${level + 1}"/><w:lvlText w:val="${level === 0 ? '%1.' : '%1.%2.'}"/></w:lvl>`).join('')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
