<script setup lang="ts">
import { computed, ref } from 'vue';
import { withBase } from 'vitepress';
import './landing.css';
import InstallPicker from './InstallPicker.vue';

const frameworks = [
	{ key: 'react', label: 'React', route: 'demo' },
	{ key: 'vue', label: 'Vue 3', route: 'demo-vue' },
	{ key: 'angular', label: 'Angular', route: 'demo-angular' },
	{ key: 'vanilla', label: 'Vanilla JS', route: 'demo-vanilla' },
	{ key: 'svelte', label: 'Svelte', route: 'demo-svelte' },
];
const active = ref(frameworks[0]);
const guest = ref('vue');
const mode = ref('solo');
const started = ref(false);
const source = computed(() =>
	withBase(
		`/${active.value.route}/${mode.value === 'collab' ? `collaboration.html?guest=${guest.value}` : ''}`,
	),
);
function selectFramework(index: number) {
	active.value = frameworks[index];
	started.value = true;
}
function moveTab(event: KeyboardEvent, index: number) {
	if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
	event.preventDefault();
	const next =
		event.key === 'Home'
			? 0
			: event.key === 'End'
				? 4
				: (index + (event.key === 'ArrowRight' ? 1 : 4)) % 5;
	selectFramework(next);
	(event.currentTarget as HTMLElement).parentElement
		?.querySelectorAll<HTMLButtonElement>('button')
		[next]?.focus();
}
</script>

<template>
	<div class="dv-landing">
		<section class="dv-hero">
			<div class="dv-hero-inner">
				<div>
					<p class="dv-kicker">WORD DOCUMENTS · TYPESCRIPT · OPEN SOURCE</p>
					<h1>Your documents.<br /><em>Your application.</em></h1>
					<p class="dv-copy">
						Open, edit and save Word documents on the web. One document engine and one editor, with
						bindings for the framework you already use.
					</p>
					<div class="dv-actions">
						<a class="dv-button" :href="withBase('/architecture')">Developer guide ↗</a
						><a class="dv-button dv-ghost" href="#live-demo">Try it live ↓</a>
					</div>
					<InstallPicker />
				</div>
				<a class="dv-preview" href="#live-demo" aria-label="Try the document editor">
					<div class="dv-framebar">
						<span class="dv-dots">● ● ●</span><span>project-brief.docx</span
						><span>Try the editor ↗</span>
					</div>
					<div class="dv-mini-ribbon">
						Home &nbsp; Insert &nbsp; Layout &nbsp; Review &nbsp; View
					</div>
					<div class="dv-paper">
						<span>PROJECT NOTES / 01</span>
						<h2>Ideas deserve<br />a clear page.</h2>
						<p>A shared workspace for your next draft.</p>
						<hr />
						<p>
							Write, refine, and work together. Keep the document at the center of your application.
						</p>
						<div class="dv-lines"></div>
					</div>
				</a>
			</div>
		</section>
		<section id="live-demo" class="dv-section">
			<p class="dv-kicker">01 / LIVE DEMO</p>
			<h2 class="dv-h2">Same editor.<br />Every framework.</h2>
			<p class="dv-copy">
				Choose your binding, open a document, and try the shared editing surface. Switch to
				collaboration to work in two editors at once.
			</p>
			<div class="dv-controls">
				<div role="tablist" aria-label="Demo framework" class="dv-tabs">
					<button
						v-for="(framework, index) in frameworks"
						:id="`framework-${framework.key}`"
						:key="framework.key"
						role="tab"
						:aria-selected="active.key === framework.key"
						aria-controls="demo-panel"
						:tabindex="active.key === framework.key ? 0 : -1"
						@click="selectFramework(index)"
						@keydown="moveTab($event, index)"
					>
						{{ framework.label }}
					</button>
				</div>
				<div class="dv-tabs" role="group" aria-label="Demo mode">
					<button
						:aria-pressed="mode === 'solo'"
						@click="
							mode = 'solo';
							started = true;
						"
					>
						Single editor</button
					><button
						:aria-pressed="mode === 'collab'"
						@click="
							mode = 'collab';
							started = true;
						"
					>
						Collaboration
					</button>
				</div>
				<label v-if="mode === 'collab'" class="dv-guest"
					>Guest binding
					<select v-model="guest">
						<option v-for="framework in frameworks" :key="framework.key" :value="framework.key">
							{{ framework.label }}
						</option>
					</select></label
				>
			</div>
			<div
				id="demo-panel"
				class="dv-live"
				role="tabpanel"
				:aria-labelledby="`framework-${active.key}`"
			>
				<div class="dv-framebar">
					<span class="dv-dots">● ● ●</span
					><span
						>{{ active.label }} ·
						{{ mode === 'solo' ? 'sample-document.docx' : 'Ada + Grace' }}</span
					><a :href="source" target="_blank" rel="noreferrer">Open full app ↗</a>
				</div>
				<iframe
					v-if="started"
					:key="source"
					:src="source"
					:title="`${active.label} ${mode === 'solo' ? 'document editor' : 'collaboration demo'}`"
					allow="clipboard-read; clipboard-write; fullscreen"
				></iframe>
				<div v-else class="dv-poster">
					<button class="dv-button" @click="started = true">Load live editor</button>
				</div>
			</div>
			<p class="dv-hint">
				{{
					mode === 'solo'
						? 'Files stay in your browser. Switching bindings starts a fresh demo.'
						: 'Two local peers share an in-memory authority. Pause delivery to test concurrent edits; this demo has no network backend.'
				}}
			</p>
		</section>
		<section class="dv-section dv-features">
			<div>
				<p class="dv-kicker">02 / ONE FOUNDATION</p>
				<h2 class="dv-h2">Built once.<br />Shared everywhere.</h2>
			</div>
			<div>
				<h3>Framework-neutral editing</h3>
				<p>
					Every binding mounts the same web component. Formatting, search, language controls and
					collaboration share one implementation.
				</p>
				<h3>Preservation with clear limits</h3>
				<p>
					Unchanged DOCX files retain their original bytes. Supported edits preserve the package;
					unsupported changes report their limits.
				</p>
				<h3>A growing Word foundation</h3>
				<p>
					Paragraphs, direct formatting, simple tables and legacy DOC text editing are available.
					Word pagination, images, lists and tracked changes remain outstanding.
				</p>
				<a :href="withBase('/parity-roadmap')">Read the support roadmap ↗</a>
			</div>
		</section>
		<section class="dv-section dv-finale">
			<p class="dv-kicker">START BUILDING</p>
			<h2 class="dv-h2">Bring your own framework.<br />Keep one document engine.</h2>
			<p class="dv-copy">
				The Word packages are not published yet. Explore the source workspace and integration guides
				while the supported editing subset grows.
			</p>
			<div class="dv-actions">
				<a class="dv-button" :href="withBase('/bindings')">Integration guides ↗</a
				><a class="dv-button dv-ghost" :href="withBase('/collaboration')">Collaboration guide ↗</a>
			</div>
		</section>
	</div>
</template>
