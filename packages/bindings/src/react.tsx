import { createElement, forwardRef, useEffect, useImperativeHandle, useRef } from 'react';
import { mountEditor, type EditorBinding, type EditorHandle, type EditorOptions } from './index';
export interface WordEditorProps extends EditorOptions {
	className?: string;
}
export const WordEditor = forwardRef<EditorHandle, WordEditorProps>(
	function WordEditor(props, ref) {
		const host = useRef<HTMLDivElement>(null);
		const binding = useRef<EditorBinding | null>(null);
		useEffect(() => {
			binding.current = mountEditor(host.current!, props);
			return () => {
				binding.current?.destroy();
				binding.current = null;
			};
		}, []);
		useEffect(() => {
			binding.current?.update(props);
		}, [props]);
		useImperativeHandle(
			ref,
			() => ({
				get element() {
					if (!binding.current) throw new Error('Editor is not mounted');
					return binding.current.element;
				},
				async load(input) {
					if (!binding.current) throw new Error('Editor is not mounted');
					await binding.current.load(input);
				},
				async save() {
					if (!binding.current) throw new Error('Editor is not mounted');
					return binding.current.save();
				},
				async download(fileName) {
					if (!binding.current) throw new Error('Editor is not mounted');
					await binding.current.download(fileName);
				},
				markClean() {
					binding.current?.markClean();
				},
				get dirty() {
					return binding.current?.dirty ?? false;
				},
			}),
			[],
		);
		return createElement('div', { ref: host, className: props.className });
	},
);
