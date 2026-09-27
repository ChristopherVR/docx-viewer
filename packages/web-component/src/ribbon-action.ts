import type { TableCommand } from './table-commands';
import type { MultilingualAction } from './multilingual-ribbon';
import type { ReviewDisplayMode } from './review-display';

/** Every command the ribbon can raise, dispatched as a `ribbon-action` event. */
export type RibbonAction =
	| {
			type: 'format';
			key: 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';
	  }
	| { type: 'history'; key: 'undo' | 'redo' }
	| { type: 'align'; value: 'left' | 'center' | 'right' | 'justify' }
	| { type: 'font'; key: 'family' | 'size' | 'color' | 'highlight'; value: string }
	| { type: 'clear' | 'table' }
	| { type: 'insertBreak'; kind: 'page' | 'column' }
	| { type: 'tableEdit'; key: TableCommand }
	| {
			type: 'page';
			key:
				| 'margin'
				| 'orientation'
				| 'columns'
				| 'numberFormat'
				| 'numberStart'
				| 'titlePage'
				| 'verticalAlign';
			value: string;
	  }
	| { type: 'sectionBreak'; kind: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' }
	| { type: 'evenOddHeaders' }
	| {
			type: 'paragraph';
			key: 'indent' | 'spacingBefore' | 'spacingAfter' | 'lineSpacing';
			value: string;
	  }
	| { type: 'zoom'; value: number }
	| { type: 'list'; key: 'bullet' | 'number' | 'increaseLevel' | 'decreaseLevel' | 'remove' }
	| { type: 'view'; value: 'draft' | 'print' }
	| { type: 'print' }
	| MultilingualAction
	| { type: 'search' }
	| {
			type: 'review';
			key:
				| 'trackChanges'
				| 'acceptOne'
				| 'rejectOne'
				| 'acceptAll'
				| 'rejectAll'
				| 'previous'
				| 'next';
	  }
	| { type: 'reviewDisplay'; value: ReviewDisplayMode }
	| { type: 'comments'; key: 'toggle' | 'add' }
	| { type: 'insertPicture' }
	| { type: 'formatPicture' }
	| { type: 'link' }
	| { type: 'characterStyle'; value: string }
	| { type: 'showHidden' }
	| { type: 'insertNote'; kind: 'footnote' | 'endnote' };
