import JSZip from 'jszip';

export const equationNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
export const fractionOmml =
	'<m:oMath><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>';
export const displayOmml =
	'<m:oMathPara><m:oMath><m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>';

/** Minimal Word package covering an inline fraction and a display superscript. */
export async function equationFixture() {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="${equationNamespace}"><w:body><w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${fractionOmml}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p><w:p>${displayOmml}</w:p><w:p><w:r><w:t>Editable paragraph</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
