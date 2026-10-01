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
	| { type: 'clear' }
	| { type: 'table'; rows?: number; columns?: number }
	| { type: 'tablePicker' }
	| { type: 'clipboard'; key: 'cut' | 'copy' | 'paste' }
	| { type: 'fontStep'; direction: 'grow' | 'shrink' }
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
				| 'verticalAlign'
				| 'lineNumbers'
				| 'size';
			value: string;
	  }
	| { type: 'sectionBreak'; kind: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' }
	| { type: 'evenOddHeaders' }
	| { type: 'pageColor'; value: string }
	| { type: 'hyphenation'; value: string }
	| { type: 'suppressLineNumbers' }
	| {
			type: 'paragraph';
			key: 'indent' | 'spacingBefore' | 'spacingAfter' | 'lineSpacing';
			value: string;
	  }
	| { type: 'zoom'; value: number }
	| {
			type: 'list';
			key:
				| 'bullet'
				| 'number'
				| 'multilevel'
				| 'outline'
				| 'headings'
				| 'article'
				| 'roman'
				| 'increaseLevel'
				| 'decreaseLevel'
				| 'remove';
	  }
	| { type: 'view'; value: 'draft' | 'print' }
	| { type: 'print' }
	| MultilingualAction
	| { type: 'search'; focus?: 'find' | 'replace' }
	| { type: 'changeCase'; value: 'sentence' | 'lower' | 'upper' | 'title' | 'toggle' }
	| { type: 'formatPainter' }
	| { type: 'ruler' }
	| { type: 'coverPage' }
	| { type: 'sort'; order: 'ascending' | 'descending' }
	| { type: 'goTo' }
	| { type: 'selectObjects' }
	| { type: 'addText'; level: 0 | 1 | 2 | 3 }
	| { type: 'blankPage' }
	| { type: 'readAloud' }
	| { type: 'gridlines' }
	| { type: 'navigation' }
	| { type: 'shading'; value: string }
	| {
			type: 'borders';
			preset:
				| 'none'
				| 'bottom'
				| 'top'
				| 'left'
				| 'right'
				| 'all'
				| 'outside'
				| 'insideH'
				| 'horizontal';
	  }
	| {
			type: 'pageNumber';
			position: 'top' | 'bottom';
			align: 'left' | 'center' | 'right';
			style?: 'plain' | 'pageOfTotal';
	  }
	| { type: 'headerFooter'; kind: 'header' | 'footer'; slot?: 'default' | 'first' | 'even' }
	| { type: 'headerFooterLink' }
	| { type: 'headerFooterClose' }
	| { type: 'headerFooterNavigate'; target: 'header' | 'footer' | 'previous' | 'next' }
	| { type: 'headerFooterPosition'; kind: 'header' | 'footer'; inches: number }
	| {
			type: 'formatDialog';
			kind: import('./format-dialogs').FormatDialogKind;
	  }
	| { type: 'indent'; side: 'left' | 'right'; inches: number }
	| { type: 'zoomFit'; mode: 'actual' | 'width' | 'page' | 'pages' }
	| { type: 'wordCount' }
	| { type: 'spelling' }
	| { type: 'insertSymbol'; value: string }
	| { type: 'dropCap'; value: 'none' | 'drop' | 'margin' }
	| { type: 'insertDateTime'; value: 'long' | 'short' | 'time' | 'datetime' }
	| { type: 'showMarks' }
	| { type: 'selectAll' }
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
	| { type: 'comments'; key: 'toggle' | 'add' | 'delete' | 'previous' | 'next' }
	| { type: 'insertPicture' }
	| { type: 'formatPicture' }
	| { type: 'link' }
	| { type: 'characterStyle'; value: string }
	| { type: 'showHidden' }
	| { type: 'thumbnails' }
	| { type: 'insertNote'; kind: 'footnote' | 'endnote' }
	| {
			type: 'toc';
			key: 'insert' | 'update' | 'remove' | 'fields' | 'figures' | 'updateFigures';
			levels?: number;
			label?: string;
	  };
