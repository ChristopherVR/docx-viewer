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
	{ key: 'solid', label: 'SolidJS', route: 'demo-solid' },
];
const active = ref(frameworks[0]);
const mode = ref<'solo' | 'collab'>('solo');
const guest = ref('vue');
const started = ref(false);
const source = computed(() =>
	withBase(
		`/${active.value.route}/${mode.value === 'collab' ? `collaboration.html?guest=${guest.value}` : ''}`,
	),
);

function selectFramework(index: number): void {
	active.value = frameworks[index];
	started.value = true;
}

function moveTab(event: KeyboardEvent, index: number): void {
	if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
	event.preventDefault();
	const next =
		event.key === 'Home'
			? 0
			: event.key === 'End'
				? frameworks.length - 1
				: (index + (event.key === 'ArrowRight' ? 1 : frameworks.length - 1)) % frameworks.length;
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
						<a class="dv-button" :href="withBase('/architecture')">Developer guide ↗</a>
						<a class="dv-button dv-ghost" href="#live-demo">Try it live ↓</a>
					</div>
					<InstallPicker />
				</div>
				<a class="dv-preview" href="#live-demo" aria-label="Try the document editor">
					<div class="dv-framebar">
						<span class="dv-dots">● ● ●</span><span>project-brief.docx</span>
						<span>Try the editor ↗</span>
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
			<h2 class="dv-h2">One editor.<br />Every framework.</h2>
			<p class="dv-copy">
				Choose a framework and try the shared editing surface. Collaboration opens two local peers
				working on one document.
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
						Single editor
					</button>
					<button
						:aria-pressed="mode === 'collab'"
						@click="
							mode = 'collab';
							started = true;
						"
					>
						Collaboration
					</button>
				</div>
				<label v-if="mode === 'collab'" class="dv-guest">
					Guest binding
					<select v-model="guest">
						<option v-for="framework in frameworks" :key="framework.key" :value="framework.key">
							{{ framework.label }}
						</option>
					</select>
				</label>
			</div>
			<div
				id="demo-panel"
				class="dv-live"
				role="tabpanel"
				:aria-labelledby="`framework-${active.key}`"
			>
				<div class="dv-framebar">
					<span class="dv-dots">● ● ●</span>
					<span
						>{{ active.label }} ·
						{{ mode === 'solo' ? 'sample-document.docx' : 'Ada + Grace' }}</span
					>
					<a :href="source" target="_blank" rel="noreferrer">Open full app ↗</a>
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
						? 'The demo document runs in your browser. Switching frameworks starts a fresh editor.'
						: 'Two local peers share an in-memory authority. Pause delivery to try concurrent edits; this demo has no network backend.'
				}}
			</p>
		</section>

		<section class="dv-section">
			<p class="dv-kicker">02 / ONE FOUNDATION</p>
			<h2 class="dv-h2">Built once.<br />Shared everywhere.</h2>
			<div class="dv-feature-grid">
				<article>
					<h3>One document model</h3>
					<p>Framework bindings adapt lifecycle and events around the same editor and model.</p>
					<a :href="withBase('/architecture')">Explore the architecture ↗</a>
				</article>
				<article>
					<h3>Careful preservation</h3>
					<p>
						Unchanged DOCX files retain their original bytes. Supported edits preserve package
						parts.
					</p>
					<a :href="withBase('/editing')">See editing behavior ↗</a>
				</article>
				<article>
					<h3>Honest format limits</h3>
					<p>Word pagination, images, lists and tracked changes are not supported yet.</p>
					<a :href="withBase('/parity-roadmap')">Read the support roadmap ↗</a>
				</article>
				<article>
					<h3>Host-owned collaboration</h3>
					<p>
						The editor exchanges validated steps; your app owns transport, identity and storage.
					</p>
					<a :href="withBase('/collaboration')">Read the collaboration guide ↗</a>
				</article>
			</div>
		</section>

		<section class="dv-section dv-agents">
			<div>
				<p class="dv-kicker">03 / YOUR APPLICATION</p>
				<h2 class="dv-h2">Keep the document<br />at the center.</h2>
				<p class="dv-copy">
					Choose a framework adapter for your interface, then use the shared model APIs for import,
					editing and saving.
				</p>
				<a class="dv-link" :href="withBase('/bindings')">Browse binding contracts ↗</a>
			</div>
			<div class="dv-code-card">
				<div class="dv-framebar"><span>editor.tsx</span><span>shared editor</span></div>
				<pre><code>import { WordEditor } from
  '@christophervr/docx-viewer/react';

&lt;WordEditor
  documentModel={model}
  onDocumentChange={setModel}
/&gt;</code></pre>
			</div>
		</section>

		<section class="dv-section dv-quickstart">
			<p class="dv-kicker">04 / QUICK START</p>
			<h2 class="dv-h2">One install.<br />The binding you need.</h2>
			<p class="dv-copy">
				Install the umbrella package once. Framework packages remain optional peers.
			</p>
			<InstallPicker />
		</section>

		<section class="dv-section dv-finale">
			<p class="dv-kicker">05 / START BUILDING</p>
			<h2 class="dv-h2">Bring your own framework.<br />Keep one document engine.</h2>
			<p class="dv-copy">
				The Word packages are not published yet. Explore the source and integration guides while the
				supported editing subset grows.
			</p>
			<div class="dv-actions">
				<a class="dv-button" :href="withBase('/bindings')">Integration guides ↗</a>
				<a class="dv-button dv-ghost" :href="withBase('/collaboration')">Collaboration guide ↗</a>
			</div>
		</section>
	</div>
</template>
