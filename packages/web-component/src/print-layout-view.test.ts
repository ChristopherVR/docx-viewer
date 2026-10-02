// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import type { DocumentModel } from 'docx-core';
import { createPrintLayoutController } from './print-layout-view';

function model(): DocumentModel {
	return {
		blocks: [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Hello Print Layout' }] }],
		page: {
			width: 300,
			height: 200,
			marginTop: 10,
			marginRight: 10,
			marginBottom: 10,
			marginLeft: 10,
		},
		warnings: [],
	};
}

describe('createPrintLayoutController', () => {
	it('starts inactive with no page status until a relayout runs', () => {
		const scrollContainer = document.createElement('div');
		const controller = createPrintLayoutController(scrollContainer, () => {});
		expect(controller.element.hidden).toBe(true);
		expect(controller.pageStatus()).toBeNull();
		controller.destroy();
	});

	it('debounces scheduleRelayout and produces a paginated render', () => {
		vi.useFakeTimers();
		try {
			const scrollContainer = document.createElement('div');
			document.body.append(scrollContainer);
			const controller = createPrintLayoutController(scrollContainer, () => {});
			controller.setActive(true);
			controller.scheduleRelayout(model());
			controller.scheduleRelayout(model()); // a second call within the debounce window should not double-render
			expect(controller.element.querySelectorAll('.dve-print-page')).toHaveLength(0);
			vi.runAllTimers();
			expect(controller.element.querySelectorAll('.dve-print-page').length).toBeGreaterThan(0);
			expect(controller.pageStatus()?.total).toBeGreaterThan(0);
			controller.destroy();
			scrollContainer.remove();
		} finally {
			vi.useRealTimers();
		}
	});

	it('invokes the cursor callback with a blockId/offset when a rendered line is clicked', () => {
		vi.useFakeTimers();
		try {
			const scrollContainer = document.createElement('div');
			document.body.append(scrollContainer);
			const onRequestCursor = vi.fn();
			const controller = createPrintLayoutController(scrollContainer, onRequestCursor);
			controller.scheduleRelayout(model());
			vi.runAllTimers();
			const line = controller.element.querySelector('.dve-print-line') as HTMLElement;
			document.body.append(controller.element);
			line.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 0 }));
			expect(onRequestCursor).toHaveBeenCalledWith('p1', expect.any(Number));
			controller.destroy();
			scrollContainer.remove();
		} finally {
			vi.useRealTimers();
		}
	});
});
