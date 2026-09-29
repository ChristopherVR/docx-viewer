import {
	ST_NumberFormat,
	type DocumentModel,
	type SectionProperties,
} from '@christophervr/docx-core';
import { setPageSize } from './page-size';
import type { RibbonAction } from './ribbon-action';
import {
	setColumns,
	setMargins,
	setOrientation,
	setPageNumbering,
	setTitlePage,
	setVerticalAlign,
} from './section-commands';

export type PageSetupKey = Extract<RibbonAction, { type: 'page' }>['key'];

const ALIGNMENTS = ['top', 'center', 'both', 'bottom'] as const;

/** The model with one Layout-tab page setting changed for section `index`. */
export function pageSetupChange(
	model: DocumentModel,
	index: number,
	section: SectionProperties,
	key: PageSetupKey,
	value: string,
): DocumentModel {
	switch (key) {
		case 'margin':
			return setMargins(model, index, value);
		case 'size':
			return setPageSize(model, index, value);
		case 'orientation':
			return setOrientation(model, index, value === 'landscape' ? 'landscape' : 'portrait');
		case 'columns':
			return setColumns(model, index, Math.max(1, Number(value) || 1));
		case 'numberFormat':
			return setPageNumbering(model, index, {
				format: ST_NumberFormat.find((item) => item === value) ?? 'decimal',
			});
		case 'numberStart':
			return setPageNumbering(model, index, { restart: value === 'restart' });
		case 'verticalAlign':
			return setVerticalAlign(model, index, ALIGNMENTS.find((item) => item === value) ?? 'top');
		case 'titlePage':
			return setTitlePage(model, index, !section.titlePage);
	}
}
